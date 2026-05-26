import os, json, re, datetime, requests
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
from dotenv import load_dotenv

load_dotenv()
app = Flask(__name__, static_folder="static")
CORS(app)

# ── Microsoft Graph OAuth ──
def get_ms_token():
    tenant = os.getenv("AZURE_TENANT_ID")
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": os.getenv("AZURE_CLIENT_ID"),
            "client_secret": os.getenv("AZURE_CLIENT_SECRET"),
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials"
        }
    )
    return resp.json().get("access_token")

def graph_get(path, token):
    r = requests.get(
        f"https://graph.microsoft.com/v1.0{path}",
        headers={"Authorization": f"Bearer {token}"}
    )
    return r.json()

# ── Parse outreach emails ──
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

# ── Routes ──
@app.route("/health")
def health():
    return jsonify({"status": "ok", "timestamp": datetime.datetime.now().isoformat()})

@app.route("/api/companies")
def get_companies():
    return jsonify(load_data())

@app.route("/api/sync", methods=["POST"])
def sync_outlook():
    try:
        token = get_ms_token()
        if not token:
            return jsonify({"error": "Could not authenticate with Microsoft Graph"}), 401

        user_email = os.getenv("USER_EMAIL", "mloebel@ftvcapital.com")

        sent = graph_get(
            f"/users/{user_email}/mailFolders/SentItems/messages"
            "?$top=500&$select=subject,toRecipients,from,sentDateTime"
            "&$orderby=sentDateTime desc",
            token
        )
        inbox = graph_get(
            f"/users/{user_email}/mailFolders/Inbox/messages"
            "?$top=500&$select=subject,from,toRecipients,receivedDateTime"
            "&$orderby=receivedDateTime desc",
            token
        )

        existing = {r["domain"]: r for r in load_data()}

        sent_messages = sent.get("value", [])
        for msg in sent_messages:
            subj = msg.get("subject", "")
            if not is_outreach(subj): continue
            for recip in msg.get("toRecipients", []):
                email = recip.get("emailAddress", {}).get("address", "")
                if not is_external(email): continue
                domain = domain_from_email(email)
                if not domain: continue
                sent_date = msg.get("sentDateTime", "")[:10]
                name = recip.get("emailAddress", {}).get("name", "")
                clean_subj = re.sub(r"^(RE:|FW:|FWD:)\s*", "", subj, flags=re.IGNORECASE).strip()
                if domain not in existing:
                    existing[domain] = {
                        "domain": domain,
                        "company": domain.split(".")[0].title(),
                        "contacts": name, "contact_emails": email,
                        "threads": [clean_subj], "num_threads": 1,
                        "total_emails": 1, "follow_ups": 0,
                        "first_contact": sent_date, "last_outreach": sent_date,
                        "responded": False, "description": "",
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
                        co["num_threads"] = len(co["threads"])
                    if not co.get("first_contact") or sent_date < co["first_contact"]:
                        co["first_contact"] = sent_date
                    if not co.get("last_outreach") or sent_date > co["last_outreach"]:
                        co["last_outreach"] = sent_date
                    co["follow_ups"] = max(0, co["total_emails"] - co.get("num_threads", 1))

        inbox_messages = inbox.get("value", [])
        for msg in inbox_messages:
            from_email = msg.get("from", {}).get("emailAddress", {}).get("address", "")
            if not is_external(from_email): continue
            domain = domain_from_email(from_email)
            if domain and domain in existing:
                existing[domain]["responded"] = True

        result = sorted(existing.values(), key=lambda r: r.get("first_contact") or "2000", reverse=True)
        save_data(result)
        return jsonify({"success": True, "total": len(result),
                        "message": f"Synced {len(sent_messages)} sent, {len(inbox_messages)} inbox. {len(result)} companies tracked."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

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
