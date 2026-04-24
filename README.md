# Personal Injury Defence Desk

A local web application for generating personal injury defence documents from a single form. Fill in the matter details once, select the documents you need, and download a ZIP of Word files ready to send.

---

## What it does

The app takes form input — matter details, parties, court information, incident facts, personnel, and billing — and generates populated `.docx` files for:

- Letters of Engagement (hourly rate and fixed fee)
- Preliminary Report
- Brief Cover Page and Index to Brief
- Memoranda of Appearance (High Court and Circuit Court)
- Notice for Particulars
- Standard Letter
- RBA Form
- Notice of Motion, Opposition, Reply
- Mediation Position Paper
- Ex Parte Application and Grounding Affidavit
- Professional Fee Note

All documents are packaged into a single ZIP download named `PIDefenceDesk_{matter_ref}_{timestamp}.zip`.

There is also a one-click **Open in Outlook** button that generates the documents and opens a pre-addressed draft in Microsoft Outlook (or macOS Mail as a fallback).

---

## Requirements

- Python 3.10 or later
- macOS (the Outlook integration uses AppleScript; all other features work on any OS)

---

## Local setup

```bash
# 1. Clone the repository
git clone https://github.com/harkintunde1995-web/pi-defencedesk.git
cd pi-defencedesk

# 2. Create and activate a virtual environment
python3 -m venv venv
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Start the server
python3 app.py
```

Then open [http://localhost:5050](http://localhost:5050) in your browser.

Alternatively, use the included start script which handles environment creation automatically:

```bash
chmod +x start.sh
./start.sh
```

---

## Project structure

```
pi-defencedesk/
├── app.py              # Flask routes: /, /generate, /send-outlook
├── generator.py        # DocumentGenerator — builds each .docx from form data
├── requirements.txt    # Python dependencies
├── start.sh            # One-command local launcher (macOS)
├── vercel.json         # Vercel deployment config
├── wsgi.py             # WSGI entry point for production
├── static/             # Static assets (CSS, images)
└── web_templates/
    └── index.html      # Single-page form UI (Tailwind CDN, Jinja2)
```

---

## Form sections

| Section | Key fields |
|---|---|
| Matter Details | File reference, document date, insurer, policy details |
| Plaintiff Details | Name, DOB, address, solicitors |
| Defendant / Insured | Client name, address, additional defendants |
| Court & Proceedings | Court/forum, record number, brief type, WRC fields |
| Personnel | Supervising partner, case handler, complaints contact |
| Fee Schedule | Hourly rates per grade, VAT rate, fixed fee |
| Incident Details | Date, time, location, injuries, PIAB |
| Preliminary Report | Executive summary, liability analysis, quantum, reserves |
| Brief / Index | Sections A–E with document lists |
| Standard Letter | Recipient, subject, body |
| Additional Particulars | Custom items for Notice for Particulars |
| Litigation Documents | Motion relief/grounds, mediation, ex parte |
| Billing / Fee Note | Time entries, disbursements, invoice preview |
| Page Spacing | Header and footer margins for pre-printed letterhead |

---

## Tech stack

| Layer | Technology |
|---|---|
| Backend | Python / Flask |
| Document generation | `python-docx` |
| Frontend | Tailwind CSS (CDN), vanilla JS |
| Templating | Jinja2 |
| Production server | Gunicorn via `wsgi.py` |
| Deployment | Vercel (serverless Python) |

---

## Deployment

The app is configured for Vercel via `vercel.json`. To deploy:

```bash
npm i -g vercel
vercel --prod
```

Note: the Outlook/AppleScript integration is macOS-only and will not function in a server deployment. The `/generate` download endpoint works on any platform.

---

## Adding a new document type

1. Add the document key and label to `DOCUMENT_LABELS` in `app.py`.
2. Implement a `generate_{key}(self, data)` method in `generator.py`.
3. Add the corresponding checkbox to the Output Builder in `web_templates/index.html`.
