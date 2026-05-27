import os, json, re, datetime, requests

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__, static_folder="static")
CORS(app, origins="*")

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

INTERNAL_DOMAINS = ["ftvcapital.com", "exchangelabs", "onmicrosoft.com", "dealcloud.intapp.com"]

def is_external(email):
    if not email: return False
    return not any(d in email.lower() for d in INTERNAL_DOMAINS)

def domain_from_email(email):
    if not email or "@" not in email: return None
    return email.split("@")[1].lower()

@app.route("/api/ingest", methods=["POST"])
def ingest():
    try:
        body = request.json
        sent_messages = body.get("sent", [])
        inbox_messages = body.get("inbox", [])

        existing = {r["domain"]: r for r in load_data()}

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
                    existing_emails = co.get("contact_emails", "")
                    if email and email not in existing_emails:
                        co["contact_emails"] = (existing_emails + ", " + email).strip(", ")

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
            "sent_processed": len(sent_messages),
            "inbox_processed": len(inbox_messages),
            "synced_at": datetime.datetime.now().isoformat()
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/health")
def health():
    return jsonify({"status": "ok", "timestamp": datetime.datetime.now().isoformat()})

@app.route("/api/companies")
def get_companies():
    return jsonify(load_data())

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
