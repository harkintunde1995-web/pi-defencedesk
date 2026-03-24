"""
PI-DefenceDesk — Personal Injury Insurance Defence Document Generator
Flask web application
"""

from flask import Flask, render_template, request, send_file, jsonify
import os
import zipfile
import io
import subprocess
from datetime import datetime
from generator import DocumentGenerator

app = Flask(__name__, template_folder="web_templates", static_folder="static")

GENERATED_DIR = os.path.join(os.path.dirname(__file__), "generated")
os.makedirs(GENERATED_DIR, exist_ok=True)

DOCUMENT_LABELS = {
    # ── Existing documents ──
    "loe_hourly":          "Letter of Engagement — Hourly Rate",
    "loe_fixed":           "Letter of Engagement — Fixed Fee",
    "preliminary_report":  "Preliminary Report",
    "brief_cover":         "Brief Cover Page",
    "index_brief":         "Index to Brief",
    "standard_letter":     "Standard Letter",
    "rba_form":            "RBA Form (Recovery of Benefits)",
    "notice_particulars":  "Notice for Particulars",
    "appearance_hc":       "Memorandum of Appearance — High Court",
    "appearance_cc":       "Memorandum of Appearance — Circuit Court",
    # ── Litigation lifecycle ──
    "motion":              "Notice of Motion",
    "opposition":          "Notice of Opposition",
    "reply":               "Reply",
    "mediation_brief":     "Mediation Position Paper",
    "ex_parte":            "Ex Parte Application + Grounding Affidavit",
    # ── Billing ──
    "fee_note":            "Professional Fee Note",
}


@app.route("/")
def index():
    today = datetime.now().strftime("%d %B %Y")
    return render_template("index.html", today=today, doc_labels=DOCUMENT_LABELS)


@app.route("/generate", methods=["POST"])
def generate():
    data = request.form.to_dict()
    selected_docs = request.form.getlist("documents")

    if not selected_docs:
        return jsonify({"error": "Please select at least one document to generate."}), 400

    generator = DocumentGenerator(data)
    generated_files = []
    errors = []

    for doc_type in selected_docs:
        try:
            path = generator.generate(doc_type)
            if path and os.path.exists(path):
                generated_files.append(path)
        except Exception as e:
            errors.append(f"{DOCUMENT_LABELS.get(doc_type, doc_type)}: {str(e)}")

    if not generated_files:
        error_msg = "No documents were generated."
        if errors:
            error_msg += " Errors: " + "; ".join(errors)
        return jsonify({"error": error_msg}), 500

    # Build ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fp in generated_files:
            zf.write(fp, os.path.basename(fp))

    zip_buffer.seek(0)

    ref = data.get("matter_ref", "documents").replace("/", "-").replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"PIDefenceDesk_{ref}_{timestamp}.zip"

    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name=zip_name,
    )


@app.route("/send-outlook", methods=["POST"])
def send_outlook():
    data = request.form.to_dict()
    selected_docs = request.form.getlist("documents")

    if not selected_docs:
        return jsonify({"error": "Please select at least one document."}), 400

    generator = DocumentGenerator(data)
    generated_files = []
    errors = []

    for doc_type in selected_docs:
        try:
            path = generator.generate(doc_type)
            if path and os.path.exists(path):
                generated_files.append(path)
        except Exception as e:
            errors.append(str(e))

    if not generated_files:
        return jsonify({"error": "No documents generated. " + "; ".join(errors)}), 500

    # Build subject + body from matter data
    matter_ref   = data.get("matter_ref", "")
    plaintiff    = data.get("plaintiff_name", "")
    recipient_to = data.get("recipient_email", data.get("insurer_contact_email", ""))
    subject      = f"{matter_ref} — {plaintiff}".strip(" — ")
    body         = (
        f"Please find attached correspondence in the above matter.\n\n"
        f"Matter Reference: {matter_ref}\n"
        f"Plaintiff: {plaintiff}\n\n"
        f"Kind regards,\n{data.get('handler_name', '')}"
    )

    # Build AppleScript — works for Microsoft Outlook on macOS
    attachments_as = "\n".join(
        [f'make new attachment with properties {{file name:POSIX file "{fp}"}} at end of attachments of newMsg'
         for fp in generated_files]
    )

    applescript = f"""
tell application "Microsoft Outlook"
    activate
    set newMsg to make new outgoing message with properties {{subject:"{subject}", plain text content:"{body}"}}
    {attachments_as}
    {"make new recipient at newMsg with properties {email address:{address:" + chr(34) + recipient_to + chr(34) + "}}" if recipient_to else ""}
    open newMsg
end tell
"""

    try:
        subprocess.run(["osascript", "-e", applescript], check=True, timeout=15)
        return jsonify({"ok": True})
    except subprocess.CalledProcessError:
        # Fallback: try macOS Mail if Outlook not installed
        applescript_mail = f"""
tell application "Mail"
    activate
    set newMsg to make new outgoing message with properties {{subject:"{subject}", content:"{body}", visible:true}}
    {"make new to recipient at newMsg with properties {address:" + chr(34) + recipient_to + chr(34) + "}" if recipient_to else ""}
    {"".join([f'tell content of newMsg to make new attachment with properties {{file name:POSIX file "{fp}"}}' for fp in generated_files])}
end tell
"""
        try:
            subprocess.run(["osascript", "-e", applescript_mail], check=True, timeout=15)
            return jsonify({"ok": True})
        except Exception as e2:
            return jsonify({"error": f"Could not open Outlook or Mail: {str(e2)}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("\n" + "=" * 55)
    print("  PI-DefenceDesk — Personal Injury Document Generator")
    print("  Open your browser at:  http://localhost:5050")
    print("=" * 55 + "\n")
    app.run(debug=False, port=5050)
