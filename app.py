import os, json, re, datetime, requests

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__, static_folder="static")
CORS(app)

# ── Load / save data ──
DATA_FILE = "data/outreach.json"
os.makedirs("data", exist_ok=True)

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE) as f:
            return json.load(f)
    return []

def save_data(d):
    with open(DATA_FILE, "w") as f:
        json.dump(d, f)

# ── Auto-sync via Claude + M365 MCP ──
OUTREACH_KEYWORDS = [
    "ftv capital intro", "ftv capital follow", "ftv intro",
    "ftv capital <>", "ftv capital /", "ftv capital in ",
    "ftv capital reconnect", "ftv capital -", "ftv capital meet",
    "ftv capital catch", "ftv capital |", "scope summit",
    "homecare 100", "ftv follow"
]
INTERNAL_DOMAINS = ["ftvcapital.com", "exchangelabs", "onmicrosoft.com"]

def is_outreach(subject):
    if not subject: return False
    s = subject.lower()
    return any(k in s for k in OUTREACH_KEYWORDS)

def is_external(email):
    if not email: return False
    return not any(d in email.lower() for d in INTERNAL_DOMAINS)

def domain_from_email(email):
    if not email or "@" not in email: return None
    return email.split("@")[1].lower()

@app.route("/api/auto-sync", methods=["POST"])
def auto_sync():
    """
    Called by Railway cron every morning.
    Uses Claude API + M365 MCP to read Outlook without Azure admin credentials.
    """
    try:
        anthropic_key = os.getenv("ANTHROPIC_API_KEY")
        if not anthropic_key:
            return jsonify({"error": "ANTHROPIC_API_KEY not set"}), 500

        user_email = os.getenv("USER_EMAIL", "mloebel@ftvcapital.com")

        # Ask Claude to fetch sent items + inbox via M365 MCP
        prompt = f"""You have access to Microsoft 365 tools. Please do the following for the mailbox {user_email}:

1. Search sent items for emails sent in the last 90 days where the subject contains any of these keywords: {', '.join(OUTREACH_KEYWORDS)}
2. For each matching sent email, extract: subject, sentDateTime, toRecipients (name + email address)
3. Search inbox for any replies received from external domains (not ftvcapital.com) in the last 90 days
4. For each inbox reply, extract: from email address, subject, receivedDateTime

Return ONLY a JSON object with this exact structure, no other text:
{{
  "sent": [
    {{"subject": "...", "sentDateTime": "YYYY-MM-DD", "toRecipients": [{{"name": "...", "email": "..."}}]}}
  ],
  "inbox": [
    {{"from_email": "...", "subject": "...", "receivedDateTime": "YYYY-MM-DD"}}
  ]
}}"""

        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": anthropic_key,
                "anthropic-version": "2023-06-01",
                "anthropic-beta": "mcp-client-2025-04-04",
                "Content-Type": "application/json"
            },
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 4000,
                "mcp_servers": [
                    {
                        "type": "url",
                        "url": "https://microsoft365.mcp.claude.com/mcp",
                        "name": "microsoft365"
                    }
                ],
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=120
        )

        if resp.status_code != 200:
            return jsonify({"error": f"Claude API error: {resp.text}"}), 500

        data = resp.json()

        # Extract JSON from Claude's response
        mail_data = None
        for block in data.get("content", []):
            if block.get("type") == "text":
                text = block["text"].strip()
                # Strip markdown code fences if present
                text = re.sub(r"```json\s*", "", text)
                text = re.sub(r"```\s*", "", text)
                try:
                    mail_data = json.loads(text)
                    break
                except:
                    # Try to find JSON object within the text
                    match = re.search(r'\{.*\}', text, re.DOTALL)
                    if match:
                        try:
                            mail_data = json.loads(match.group())
                            break
                        except:
                            pass

        if not mail_data:
            return jsonify({"error": "Could not parse mail data from Claude response", "raw": str(data)}), 500

        # Process sent emails
        existing = {r["domain"]: r for r in load_data()}
        sent_messages = mail_data.get("sent", [])
        inbox_messages = mail_data.get("inbox", [])

        for msg in sent_messages:
            subj = msg.get("subject", "")
            sent_date = msg.get("sentDateTime", "")[:10]
            clean_subj = re.sub(r"^(RE:|FW:|FWD:)\s*", "", subj, flags=re.IGNORECASE).strip()

            for recip in msg.get("toRecipients", []):
                email = recip.get("email", "")
                name = recip.get("name", "")
                if not is_external(email): continue
                domain = domain_from_email(email)
                if not domain: continue

                if domain not in existing:
                    existing[domain] = {
                        "domain": domain,
                        "company": domain.split(".")[0].title(),
                        "contacts": name,
                        "contact_emails": email,
                        "threads": [clean_subj],
                        "num_threads": 1,
                        "total_emails": 1,
                        "follow_ups": 0,
                        "first_contact": sent_date,
                        "last_outreach": sent_date,
                        "responded": False,
                        "description": "",
                        "hq_city": None, "hq_state": None, "hq_country": None,
                        "ceo": None, "company_linkedin": None,
                        "employees": None, "emp_growth_1yr": None,
                        "website": f"https://{domain}"
                    }
                else:
                    co = existing[domain]
                    co["total_emails"] = co.get("total_emails", 0) + 1
                    if clean_subj not in co.get("threads", []):
                        co.setdefault("threads", []).append(clean_subj)
                    co["num_threads"] = len(co.get("threads", []))
                    if not co.get("first_contact") or sent_date < co["first_contact"]:
                        co["first_contact"] = sent_date
                    if not co.get("last_outreach") or sent_date > co["last_outreach"]:
                        co["last_outreach"] = sent_date
                    co["follow_ups"] = max(0, co["total_emails"] - co.get("num_threads", 1))
                    # Add new contact email if not already tracked
                    existing_emails = co.get("contact_emails", "")
                    if email and email not in existing_emails:
                        co["contact_emails"] = (existing_emails + ", " + email).strip(", ")

        # Process inbox replies
        for msg in inbox_messages:
            from_email = msg.get("from_email", "")
            if not is_external(from_email): continue
            domain = domain_from_email(from_email)
            if domain and domain in existing:
                existing[domain]["responded"] = True

        result = sorted(existing.values(), key=lambda r: r.get("first_contact") or "2000", reverse=True)
        save_data(result)

        return jsonify({
            "success": True,
            "total": len(result),
            "new_sent_processed": len(sent_messages),
            "inbox_processed": len(inbox_messages),
            "message": f"Auto-sync complete. {len(result)} companies tracked.",
            "synced_at": datetime.datetime.now().isoformat()
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/sync-status")
def sync_status():
    """Returns last sync time and company count."""
    data = load_data()
    return jsonify({
        "total_companies": len(data),
        "last_synced": datetime.datetime.now().isoformat()
    })


# ── Keep existing manual sync for backup (now uses Claude too) ──
@app.route("/api/sync", methods=["POST"])
def sync_outlook():
    """Manual sync — delegates to auto_sync logic."""
    return auto_sync()


# ── Health ──
@app.route("/health")
def health():
    return jsonify({"status": "ok", "timestamp": datetime.datetime.now().isoformat()})

@app.route("/api/companies")
def get_companies():
    return jsonify(load_data())

# ── AI Chat ──
@app.route("/api/chat", methods=["POST"])
def chat():
    body = request.json
    messages = body.get("messages", [])
    companies = load_data()

    lines = []
    for r in companies:
        hq = ", ".join(filter(None, [r.get("hq_city"), r.get("hq_state") or r.get("hq_country")]))
        lines.append("|".join([
            r["company"],
            (r.get("description") or "")[:60],
            hq,
            r.get("ceo") or "",
            str(r.get("employees") or ""),
            str(r.get("emp_growth_1yr") or "") + "%",
            r.get("first_contact") or "",
            r.get("last_outreach") or "",
            str(r.get("follow_ups", 0)) + "fu",
            str(r.get("total_emails", 0)) + "emails",
            "responded" if r.get("responded") else "no_response",
            r.get("domain", "")
        ]))

    system = f"""You are an AI assistant in Matt Loebel's FTV Capital outreach tracker.
Today: {datetime.date.today().isoformat()}

DATA (Company|Desc|HQ|CEO|Employees|Growth|FirstContact|LastOutreach|FUs|Emails|Status|Domain):
{chr(10).join(lines)}

Answer questions about outreach data concisely. Use HTML tables (class="ai-table") for lists."""

    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": os.getenv("ANTHROPIC_API_KEY"),
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
        },
        json={"model": "claude-sonnet-4-20250514", "max_tokens": 1500,
              "system": system, "messages": messages}
    )
    return jsonify(resp.json())

@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve(path):
    return send_from_directory("static", "index.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
