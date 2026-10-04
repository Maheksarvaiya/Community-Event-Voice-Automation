📞 Community Event Voice Automation

Automated AI voice-call outreach for community events. The system reads members and events from an Excel sheet, submits a call cohort to a Sarvam AI voice agent, and receives call results through a Flask webhook that writes attendance responses straight back into Excel.

Excel in → AI voice calls → Webhook → Excel updated. No manual follow-ups.

✨ Features
Excel-driven workflow – Events and Members live in a single .xlsx workbook (Events and Members sheets).
Cohort builder – Selects eligible members, validates phone numbers (E.164-style Indian format, +91XXXXXXXXXX) and attaches the matching event details.
Personalised calls – Passes customer_name, event_name, event_date, event_time and event_location to the voice agent as app variables.
Duplicate-call protection – Skips any member already called today, so an accidental re-run won't spam people.
Dry-run safety switch – DRY_RUN = True prepares and prints the cohort without making any API call.
Robust date parsing – Handles DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD and datetime cells.
Webhook receiver – Flask endpoint that updates call status, attendance response and timestamp per member.
Clean attendance states – CONFIRMED, DECLINED, MAYBE, CALLBACK, UNCLEAR.
Fail-fast errors – Clear messages for a missing API key, Excel file or sheet.
🏗️ How It Works
┌──────────────────┐     ┌────────────────┐     ┌─────────────────┐
│  community_event │     │  send_cohort.py│     │  Sarvam Voice   │
│  _automation.xlsx├────►│  validate +    ├────►│  Agent (calls   │
│  Events/Members  │     │  build cohort  │     │  members)       │
└────────▲─────────┘     └────────────────┘     └────────┬────────┘
         │                                               │ results
         │            ┌────────────────┐                 │
         └────────────┤   webhook.py   │◄────────────────┘
          update row  │  (Flask /POST) │   POST /webhook
                      └────────────────┘
send_cohort.py loads the workbook, filters eligible members and validates phone numbers.
It builds the cohort (using the member ID as the stable user_identifier) and submits it to Sarvam.
After each call, Sarvam POSTs the result to /webhook.
webhook.py finds the member by ID and writes the call status, attendance response and timestamp back to Excel.
📁 Project Structure
product-ai-extractor/
├── send_cohort.py                  # Builds and submits the call cohort
├── webhook.py                      # Flask webhook that updates Excel
├── community_event_automation.xlsx # Events + Members data (keep private)
├── requirements.txt
├── .env                            # API keys (never commit)
└── .gitignore
🚀 Getting Started
1. Clone and install
bash
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>

python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
2. Configure environment

Create a .env file:

env
VOICE_AGENTS_API_KEY=your_sarvam_api_key
3. Prepare the Excel file

community_event_automation.xlsx needs two sheets:

Sheet	Purpose
Events	Event name, date, time, location
Members	Member ID, name, phone, event link, call status, attendance response, last call time
4. Run a dry run first

In send_cohort.py:

python
DRY_RUN = True   # prepare cohort only, NO API calls
bash
python send_cohort.py

You'll see the list of eligible members (phone numbers masked to the last 4 digits) and a skipped-members summary.

5. Start the webhook
bash
python webhook.py

The server runs on http://0.0.0.0:5000. To receive Sarvam callbacks locally, expose it with a tunnel such as ngrok:

bash
ngrok http 5000

Set the public URL (https://<id>.ngrok.app/webhook) as the webhook in your Sarvam campaign.

6. Go live
python
DRY_RUN = False  # actually submit the cohort to Sarvam
bash
python send_cohort.py
🔌 Webhook Payload

POST /webhook expects JSON like:

json
{
  "user_identifier": "MEM-001",
  "completion_status": "completed",
  "connectivity_status": "connected",
  "output_agent_variables": {
    "attendance_response": "CONFIRMED"
  }
}
Field	Description
user_identifier	Member ID used to locate the row in Excel
connectivity_status	connected → call COMPLETED, otherwise FAILED
attendance_response	One of CONFIRMED, DECLINED, MAYBE, CALLBACK, UNCLEAR

Responses: 200 on success, 400 for invalid JSON or an unknown member.

🛠️ Tech Stack
Python 3.12
Flask – webhook server
openpyxl – Excel read/write
requests – Sarvam API calls
python-dotenv – environment configuration
Sarvam AI Voice Agents – conversational calling
🔒 Security Notes
Never commit .env or real member data. Keep community_event_automation.xlsx out of the repo (or commit a sanitised sample).
Run with debug=False and put the webhook behind HTTPS in production.
Consider validating a shared secret/signature on incoming webhook requests.
🗺️ Roadmap
 Webhook authentication (shared secret)
 Retry logic for failed calls
 Move from Excel to a database / Odoo integration
 Dashboard for attendance analytics
 Multi-language call scripts
👤 Author

Mahek — Python & Odoo developer | AI Automation Connect on LinkedIn · GitHub

📄 License

MIT — see LICENSE for details.
