"""
DWF (Ireland) LLP — Automated Document Generator
Generates: Letters of Engagement, Preliminary Reports, Brief Covers,
           Index to Brief, RBA Forms, Notice for Particulars, Appearances
"""

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
from datetime import datetime

OUTPUT_DIR = "/tmp/generated"


class DocumentGenerator:

    def __init__(self, data):
        self.data = data
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    def get(self, key, default="[TBC]"):
        val = self.data.get(key, "")
        return val.strip() if val and val.strip() else default

    def generate(self, doc_type):
        generators = {
            "loe_hourly": self.gen_loe_hourly,
            "loe_fixed": self.gen_loe_fixed,
            "preliminary_report": self.gen_preliminary_report,
            "brief_cover": self.gen_brief_cover,
            "index_brief": self.gen_index_brief,
            "standard_letter": self.gen_standard_letter,
            "rba_form": self.gen_rba_form,
            "notice_particulars": self.gen_notice_particulars,
            "appearance_hc": self.gen_appearance_hc,
            "appearance_cc": self.gen_appearance_cc,
        }
        fn = generators.get(doc_type)
        return fn() if fn else None

    # ─────────────────────────────────────────────────────────────
    # HELPERS
    # ─────────────────────────────────────────────────────────────

    def _new_doc(self):
        doc = Document()
        try:
            top = Cm(float(self.data.get("header_space") or 2.5))
        except (ValueError, TypeError):
            top = Cm(2.5)
        try:
            bottom = Cm(float(self.data.get("footer_space") or 2.5))
        except (ValueError, TypeError):
            bottom = Cm(2.5)
        for section in doc.sections:
            section.top_margin    = top
            section.bottom_margin = bottom
            section.left_margin   = Cm(2.5)
            section.right_margin  = Cm(2.5)
        return doc

    def _para(self, doc, text="", bold=False, italic=False, size=11,
              align=WD_ALIGN_PARAGRAPH.LEFT, sp_before=0, sp_after=6,
              color=None, font="Calibri", underline=False):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(sp_before)
        p.paragraph_format.space_after = Pt(sp_after)
        if text:
            run = p.add_run(text)
            run.font.name = font
            run.font.size = Pt(size)
            run.bold = bold
            run.italic = italic
            run.underline = underline
            if color:
                run.font.color.rgb = RGBColor(*color)
        return p

    def _heading(self, doc, text, size=11, underline=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(size)
        run.bold = True
        run.underline = underline
        return p

    def _cell_text(self, cell, text, bold=False, size=10,
                   align=WD_ALIGN_PARAGRAPH.LEFT, font="Calibri"):
        for p in cell.paragraphs:
            for run in p.runs:
                run.text = ""
        p = cell.paragraphs[0]
        p.clear()
        p.alignment = align
        run = p.add_run(str(text))
        run.font.name = font
        run.font.size = Pt(size)
        run.bold = bold
        return p

    def _cell_bg(self, cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        tcPr.append(shd)

    def _remove_table_borders(self, tbl):
        """Make a table borderless (clean letterhead look)."""
        for row in tbl.rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcBorders = OxmlElement("w:tcBorders")
                for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                    border = OxmlElement(f"w:{side}")
                    border.set(qn("w:val"), "none")
                    tcBorders.append(border)
                tcPr.append(tcBorders)

    def _letterhead_table(self, doc):
        """Standard DWF 2-column letterhead info table."""
        client_addr = self.get("client_address", "")
        recipient_block = "\n".join(
            filter(None, [self.get("client_name"), client_addr])
        )
        tbl = doc.add_table(rows=4, cols=2)
        self._remove_table_borders(tbl)

        self._cell_text(tbl.rows[0].cells[0], recipient_block, size=10)
        self._cell_text(tbl.rows[0].cells[1], "Private & Confidential", bold=True, size=10)

        self._cell_text(tbl.rows[1].cells[0], "", size=10)
        self._cell_text(tbl.rows[1].cells[1],
            f"Our Ref: {self.get('matter_ref')}\nPlease quote this when replying", size=10)

        self._cell_text(tbl.rows[2].cells[0], "By post", size=10)
        self._cell_text(tbl.rows[2].cells[1],
            f"Date: {self.get('matter_date', datetime.now().strftime('%d %B %Y'))}", size=10)

        self._cell_text(tbl.rows[3].cells[0], "", size=10)
        self._cell_text(tbl.rows[3].cells[1],
            f"Please ask for: {self.get('handler_name')}\nE-mail: {self.get('handler_email')}",
            size=10)
        return tbl

    def _ext_letterhead_table(self, doc, name="", firm="", address=""):
        """Letterhead table for letters going to external parties."""
        block = "\n".join(filter(None, [name, firm, address]))
        tbl = doc.add_table(rows=4, cols=2)
        self._remove_table_borders(tbl)

        self._cell_text(tbl.rows[0].cells[0], block, size=10)
        self._cell_text(tbl.rows[0].cells[1], "Private & Confidential", bold=True, size=10)

        self._cell_text(tbl.rows[1].cells[0], "", size=10)
        self._cell_text(tbl.rows[1].cells[1],
            f"Our Ref: {self.get('matter_ref')}\nPlease quote this when replying", size=10)

        self._cell_text(tbl.rows[2].cells[0], "By post", size=10)
        self._cell_text(tbl.rows[2].cells[1],
            f"Date: {self.get('matter_date', datetime.now().strftime('%d %B %Y'))}", size=10)

        self._cell_text(tbl.rows[3].cells[0], "", size=10)
        self._cell_text(tbl.rows[3].cells[1],
            f"Please ask for: {self.get('handler_name')}\nE-mail: {self.get('handler_email')}",
            size=10)
        return tbl

    def _matter_title(self, doc):
        """Plaintiff v Defendant + court record line."""
        self._para(doc,
            f"{self.get('plaintiff_name')} v {self.get('client_name')}",
            bold=True, size=11, sp_before=6, sp_after=3)
        record = self.get("record_no", "")
        court = self.get("court_type", "High Court")
        if record and record != "[TBC]":
            self._para(doc, f"{court} Record Number {record}", size=11, sp_after=6)

    def _court_caption(self, doc, font="Times New Roman"):
        """Full court caption block (High Court style) for pleadings."""
        court = self.get("court_type", "High Court")
        record_no = self.get("record_no", "")
        plaintiff = self.get("plaintiff_name")
        client = self.get("client_name")
        add_defs = self.get("additional_defendants", "")

        if "wrc" in court.lower() or "workplace" in court.lower():
            court_display = "WORKPLACE RELATIONS COMMISSION"
        elif "circuit" in court.lower():
            court_display = "THE CIRCUIT COURT"
        elif "district" in court.lower():
            court_display = "THE DISTRICT COURT"
        else:
            court_display = "THE HIGH COURT"

        self._para(doc, court_display, bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3, font=font)
        if record_no and record_no != "[TBC]":
            self._para(doc, f"Record No. {record_no}", bold=True, size=12,
                       align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=6, font=font)
        self._para(doc, "BETWEEN:", bold=True, size=12,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3, font=font)
        self._para(doc, plaintiff.upper(), bold=True, size=12,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3, font=font)
        self._para(doc, "PLAINTIFF", size=11,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3, font=font)
        self._para(doc, "-AND-", bold=True, size=12,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3, font=font)
        self._para(doc, client.upper(), bold=True, size=12,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3, font=font)
        if add_defs and add_defs != "[TBC]":
            self._para(doc, add_defs.upper(), bold=True, size=12,
                       align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3, font=font)
        self._para(doc, "DEFENDANT(S)", size=11,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12, font=font)

    def _dwf_sig_block(self, doc, font="Calibri"):
        """Standard DWF signature block."""
        partner = self.get("partner_name", "")
        handler = self.get("handler_name", "")
        handler_title = self.get("handler_title", "Solicitor")
        if partner and partner != "[TBC]":
            self._para(doc, partner, bold=True, size=11, font=font)
            self._para(doc, "Partner", size=11, font=font)
        elif handler and handler != "[TBC]":
            self._para(doc, handler, bold=True, size=11, font=font)
            self._para(doc, handler_title, size=11, font=font)
        self._para(doc, "DWF (Ireland) LLP", size=11, font=font)

    def _dwf_address_block(self, doc, font="Calibri"):
        """DWF address block for pleadings."""
        ref = self.get("matter_ref", "")
        defended = self.get("defended_party", "Defendant")
        self._para(doc, "DWF (Ireland) LLP", bold=True, size=11, font=font)
        self._para(doc, "The Lennox", size=11, font=font, sp_after=0)
        self._para(doc, "50 Richmond Street South", size=11, font=font, sp_after=0)
        self._para(doc, "Saint Kevin's", size=11, font=font, sp_after=0)
        self._para(doc, "Dublin 2", size=11, font=font, sp_after=3)
        self._para(doc, f"Solicitors for the {defended}", italic=True, size=11, font=font)
        if ref and ref != "[TBC]":
            self._para(doc, f"(Ref: {ref})", size=11, font=font)

    def _save(self, doc, filename):
        path = os.path.join(OUTPUT_DIR, filename)
        doc.save(path)
        return path

    def _safe_filename(self, ref, name, prefix, ext="docx"):
        ref_clean = ref.replace("/", "-").replace(" ", "_")
        name_clean = name.replace(" ", "_")[:20]
        return f"{prefix}_{ref_clean}_{name_clean}.{ext}"

    # ─────────────────────────────────────────────────────────────
    # 1. LETTER OF ENGAGEMENT — HOURLY
    # ─────────────────────────────────────────────────────────────

    def gen_loe_hourly(self):
        doc = self._new_doc()
        self._letterhead_table(doc)
        self._para(doc, "")

        self._para(doc, "Dear Sirs", size=11)
        self._matter_title(doc)

        insurer = self.get("insurer_name", "the Insurers")
        partner = self.get("partner_name", "the Partner")
        handler = self.get("handler_name")
        handler_title = self.get("handler_title", "Solicitor")
        handler_email = self.get("handler_email")
        handler_phone = self.get("handler_phone", "01 790 9400")
        policy_no = self.get("policy_no", "[POLICY NUMBER]")
        p_from = self.get("policy_period_from", "[DATE]")
        p_to = self.get("policy_period_to", "[DATE]")
        excess = self.get("policy_excess", "[EXCESS AMOUNT]")
        complaints_partner = self.get("complaints_partner", "Chris Wheeler")
        complaints_phone = self.get("complaints_phone", "01 790 9400")
        complaints_email = self.get("complaints_email", "Chris.Wheeler2@dwf.law")

        self._para(doc,
            f"We confirm that DWF have been appointed by your insurers, '{insurer}' (\"Insurers\") "
            f"to act on behalf of you in relation to the above matter. The purpose of this letter is "
            f"to explain the basis upon which DWF will act and also to provide your firm with important "
            f"initial information regarding the handling of this matter.",
            size=11, sp_after=6)

        self._heading(doc, "DWF's retainer")
        self._para(doc,
            "DWF have been retained to investigate and defend the claim and report to Insurers on "
            "liability and strategy.", size=11, sp_after=6)
        self._para(doc,
            f"DWF are instructed subject to the terms and conditions of your firm's Policy Number "
            f"{policy_no} for the period {p_from} to {p_to} (\"The Policy\") which also governs the "
            f"relationship between you and your Insurers. These terms and conditions, so far as "
            f"relevant, are incorporated into DWF's contract with you. You should therefore read "
            f"the Policy carefully.", size=11, sp_after=6)
        self._para(doc,
            "Neither this letter nor our acting for you should be taken by you as terminating or "
            "waiving your Insurer's reservation of their position, where relevant, or as terminating "
            "or waiving your Insurer's right to withdraw or refuse to grant an indemnity to you, "
            "whether your Insurer has communicated their position on indemnity to you or not.",
            size=11, sp_after=6)

        self._heading(doc, "Client")
        self._para(doc,
            "Your Insurer will determine how this Matter should be handled and we will act for both "
            "you and your Insurer upon your Insurer's instructions, in accordance with our retainer "
            "with your Insurer, without specific confirmation from you, on the basis that your Insurer "
            "has your authority to give such instructions. Any information or documents which you "
            "provide to us in connection with this Matter will also be made available to your Insurer.",
            size=11, sp_after=6)
        self._para(doc,
            "If we decide to or are required to stop acting for you, we will provide you with as much "
            "notice as possible of this and, where possible, explain the reasons to you and any "
            "possible options for you to pursue the Matter.", size=11, sp_after=6)
        self._para(doc,
            "Unless we agree in writing to the contrary, the advice provided and work carried out by "
            "us in relation to this Matter is intended to be relied upon by you and your Insurer only "
            "and by no other party.", size=11, sp_after=6)

        self._heading(doc, "Personnel")
        first = handler.split()[0] if handler and handler != "[TBC]" else "the case handler"
        self._para(doc,
            f"The Supervising Partner is {partner} and the case will be handled by {handler}, "
            f"{handler_title}. If you have any queries about your Matter or you would like further "
            f"information about any aspect of your Matter then you can contact {handler} by telephone "
            f"on {handler_phone} or alternatively, you can send an e-mail to {handler_email}. "
            f"If {first} is not available then her secretary or another member of the team should be "
            f"able to assist you. We will notify you if the personnel handling your Matter changes.",
            size=11, sp_after=6)

        self._heading(doc, "Your Involvement")
        self._para(doc,
            "Although your Insurers are dealing with the Matter on your behalf, and, once proceedings "
            "are issued, you are subject to the rules of the court. You also have a duty to assist your "
            "Insurer as a condition of having insurance cover.", size=11, sp_after=6)
        self._para(doc,
            "You are required to give your full and prompt co-operation to any requests for information "
            "and documentation throughout the course of the litigation in respect of this Matter. There "
            "is a very real danger that, no matter the strengths of your Defence, it will be unsuccessful "
            "or dismissed if there is a failure to abide by the rules of the court or any court orders.",
            size=11, sp_after=6)

        self._heading(doc, "Outline of work to be done for your case")
        self._para(doc,
            "In order to keep you informed about your case, we include here an outline of the work to "
            "be done in respect of each stage of the litigation process:", size=11, sp_after=4)
        for item in [
            "Pre-Litigation investigations (including review of insured's file of papers, Injuries Board "
            "documentation and proceedings, engage with insured, interview appropriate witnesses and "
            "instruct witnesses where necessary). Consideration also to be given to joinder of other "
            "parties to the proceedings.",
            "Investigate possibility of compromising litigation;",
            "Where necessary, consider alternative dispute resolution including mediation;",
            "Where necessary, instruct Counsel to draft Notice for Particulars, Defence, request for "
            "Discovery and advise on liability and quantum;",
            "Instruct appropriate expert witnesses (medical, actuarial and/or other);",
            "Where compromise not possible, obtain Advice on Proofs and commence preparation for trial;",
            "Trial",
        ]:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(item)
            r.font.name = "Calibri"
            r.font.size = Pt(11)

        self._heading(doc, "Disclosure")
        self._para(doc,
            "You should be aware that Insurers will be given full access to all information which comes "
            "into DWF's possession relating to this matter, including all documents or copies thereof "
            "arising from or in connection with the underlying facts. Insurers may be given copies of "
            "correspondence between DWF and others and opinions or advice from counsel, as well as "
            "notes of conversations and meetings or consultations with you or counsel.",
            size=11, sp_after=6)
        self._para(doc,
            "We must draw to your attention that the duty of standard disclosure requires parties to a "
            "claim to disclose relevant documents, including those stored in electronic form. You must "
            "therefore give consideration to information which may be stored and/or associated with "
            "electronic documents, including email and other electronic communication, word processed "
            "documents and databases. You have an immediate and ongoing duty to preserve all document "
            "and electronic data which may be relevant to the matter.",
            size=11, sp_after=6)

        self._heading(doc, "Defence costs")
        self._para(doc,
            "DWF are entitled to submit invoices for professional fees and other costs to you as a "
            "client of the firm. In practice, however, where our fees are to be paid by Insurers we "
            "will send our invoices directly to them. I must advise you, however, that in the unlikely "
            "event that Insurers are unable to pay DWF's invoices you remain liable for all fees and "
            "costs incurred.", size=11, sp_after=6)
        self._para(doc,
            "As we wish you to be informed about our charges, we set out the basis on which our legal "
            "costs will be calculated. This is as follows:", size=11, sp_after=6)

        # Fee table
        fee_tbl = doc.add_table(rows=6, cols=2)
        fee_tbl.style = "Table Grid"
        headers = ["Fee Earner", "Hourly Charge Out Rate"]
        for i, h in enumerate(headers):
            self._cell_text(fee_tbl.rows[0].cells[i], h, bold=True, size=10)
            self._cell_bg(fee_tbl.rows[0].cells[i], "D9D9D9")
        for i, (role, key, default) in enumerate([
            ("Partner", "rate_partner", "300"),
            ("Senior Associate", "rate_senior_associate", "250"),
            ("Associate", "rate_associate", "250"),
            ("Solicitor", "rate_solicitor", "250"),
            ("Paralegal", "rate_paralegal", "150"),
        ]):
            row = fee_tbl.rows[i + 1]
            self._cell_text(row.cells[0], role, size=10)
            self._cell_text(row.cells[1], f"€{self.get(key, default)}", size=10)

        self._para(doc, "")
        vat = self.get("vat_rate", "23")
        self._para(doc,
            f"The amount of value-added tax to be charged in respect of the above is {vat}%. "
            f"The above charges have been calculated by reference to the following matters:",
            size=11, sp_after=4)
        for item in [
            "the complexity and novelty of the issues involved in the legal work;",
            "the skill or specialised knowledge relevant to the matter which the legal practitioner "
            "has applied to the matter;",
            "the time and labour that the legal practitioner has reasonably expended on the matter;",
            "the urgency attached to the matter by the client;",
            "the place and circumstances in which the matter was transacted;",
            "the number, importance and complexity of the documents that the legal practitioner was "
            "required to draft, prepare or examine;",
            "whether or not the legal practitioner necessarily undertook research or investigative work;",
            "the use and costs of expert witnesses or other expertise engaged by the legal practitioner.",
        ]:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(item)
            r.font.name = "Calibri"
            r.font.size = Pt(11)

        self._para(doc, "")
        self._para(doc,
            "We should be obliged if you would confirm whether you are registered for VAT and if so, "
            "DWF will invoice you in respect of the VAT on our fees (and any Counsel's fees) in "
            "accordance with our usual practice and VAT legislation.", size=11, sp_after=6)

        self._heading(doc, "Your excess")
        self._para(doc,
            f"I understand that the excess payable under the Policy is €{excess} which is payable "
            f"in accordance with the terms of the Policy.", size=11, sp_after=6)

        self._heading(doc, "Liability for the Claimant's costs")
        self._para(doc,
            "I will provide you with my recommendations in relation to future strategy and settlement "
            "options (where appropriate) once I have had the opportunity to consider all the relevant "
            "documentation. I am also professionally obliged to draw to your attention that if this "
            "matter proceeds to trial and the other party is successful, the Court may make an adverse "
            "costs award against you. You will be liable for those costs if that is the case, subject "
            "to any claim for indemnity you may have under the terms of the Policy.",
            size=11, sp_after=6)

        self._heading(doc, "Complaints Procedure")
        self._para(doc,
            f"DWF are committed to high quality legal advice and client care. If you are unhappy about "
            f"any aspect of the service you receive, please contact me. If you are not satisfied with "
            f"my response, you should contact our Partner, {complaints_partner} (telephone: "
            f"{complaints_phone}; email: {complaints_email}). We also attach a copy of our complaints "
            f"procedure for your attention.", size=11, sp_after=6)

        self._heading(doc, "Next Steps")
        self._para(doc,
            "Finally, we would be grateful if you would sign and date a copy of this letter below "
            "and return it to us at your earliest convenience. In the event that we do not receive a "
            "countersigned copy of this letter from you, we will assume, if you continue to instruct "
            "us, that you have accepted and agreed to the terms of engagement as set out in this letter.",
            size=11, sp_after=6)

        self._para(doc, "")
        self._para(doc, "Yours sincerely,", size=11)
        self._para(doc, "")
        self._para(doc, "")
        self._dwf_sig_block(doc)
        self._para(doc, "")
        date = self.get("matter_date", datetime.now().strftime("%d %B %Y"))
        self._para(doc,
            f"We acknowledge receipt of this letter dated {date} and agree to its terms.",
            size=11, sp_after=6)
        self._para(doc, f"Signed on behalf of {self.get('client_name')}", size=11, sp_after=20)
        self._para(doc, "_" * 45, size=11)

        ref = self.get("matter_ref", "LOE")
        return self._save(doc, self._safe_filename(ref, self.get("client_name", "client"), "LOE_Hourly"))

    # ─────────────────────────────────────────────────────────────
    # 2. LETTER OF ENGAGEMENT — FIXED FEE
    # ─────────────────────────────────────────────────────────────

    def gen_loe_fixed(self):
        doc = self._new_doc()
        self._letterhead_table(doc)
        self._para(doc, "")

        self._para(doc, "Dear Sirs", size=11)
        self._matter_title(doc)

        insurer = self.get("insurer_name", "the Insurers")
        partner = self.get("partner_name", "the Partner")
        handler = self.get("handler_name")
        handler_title = self.get("handler_title", "Solicitor")
        handler_email = self.get("handler_email")
        handler_phone = self.get("handler_phone", "01 790 9400")
        policy_no = self.get("policy_no", "[POLICY NUMBER]")
        p_from = self.get("policy_period_from", "[DATE]")
        p_to = self.get("policy_period_to", "[DATE]")
        fixed_fee = self.get("fixed_fee", "[FEE AMOUNT]")
        excess = self.get("policy_excess", "[EXCESS AMOUNT]")
        vat = self.get("vat_rate", "23")
        complaints_partner = self.get("complaints_partner", "Chris Wheeler")
        complaints_phone = self.get("complaints_phone", "01 790 9400")
        complaints_email = self.get("complaints_email", "Chris.Wheeler2@dwf.law")

        self._para(doc,
            f"We confirm that DWF have been appointed by your insurers, '{insurer}' (\"Insurers\") "
            f"to act on behalf of you in relation to the above matter.", size=11, sp_after=6)

        self._heading(doc, "DWF's retainer")
        self._para(doc,
            "DWF have been retained to investigate and defend the claim and report to Insurers on "
            "liability and strategy.", size=11, sp_after=6)
        self._para(doc,
            f"DWF are instructed subject to the terms and conditions of your firm's Policy Number "
            f"{policy_no} for the period {p_from} to {p_to} (\"The Policy\").",
            size=11, sp_after=6)

        self._heading(doc, "Fixed Fee Arrangement")
        self._para(doc,
            f"In accordance with our agreement with your Insurers, DWF will carry out all work "
            f"required to defend this claim on a fixed fee basis. The agreed fixed fee for this "
            f"matter is €{fixed_fee} plus VAT at {vat}%.", size=11, sp_after=6)
        self._para(doc,
            "This fixed fee covers all work required to bring this matter to resolution, including "
            "all correspondence, preparation of pleadings, briefing of counsel, attendance at "
            "consultations, and, where required, preparation for and attendance at trial.",
            size=11, sp_after=6)

        self._heading(doc, "Personnel")
        self._para(doc,
            f"The Supervising Partner is {partner} and the case will be handled by {handler}, "
            f"{handler_title}. Contact: {handler_email} or {handler_phone}.", size=11, sp_after=6)

        self._heading(doc, "Your excess")
        self._para(doc,
            f"The excess payable under the Policy is €{excess} which is payable in accordance "
            f"with the terms of the Policy.", size=11, sp_after=6)

        self._heading(doc, "Complaints Procedure")
        self._para(doc,
            f"DWF are committed to high quality legal advice and client care. If you are unhappy "
            f"about any aspect of the service you receive, please contact our Partner, "
            f"{complaints_partner} (telephone: {complaints_phone}; email: {complaints_email}).",
            size=11, sp_after=6)

        self._heading(doc, "Next Steps")
        self._para(doc,
            "Please sign and return a copy of this letter confirming your agreement to the terms "
            "of engagement.", size=11, sp_after=12)

        self._para(doc, "Yours sincerely,", size=11)
        self._para(doc, "")
        self._para(doc, "")
        self._dwf_sig_block(doc)
        self._para(doc, "")
        date = self.get("matter_date", datetime.now().strftime("%d %B %Y"))
        self._para(doc,
            f"We acknowledge receipt of this letter dated {date} and agree to its terms.",
            size=11, sp_after=20)
        self._para(doc, "_" * 45, size=11)

        ref = self.get("matter_ref", "LOE")
        return self._save(doc, self._safe_filename(ref, self.get("client_name", "client"), "LOE_Fixed"))

    # ─────────────────────────────────────────────────────────────
    # 3. PRELIMINARY REPORT
    # ─────────────────────────────────────────────────────────────

    def gen_preliminary_report(self):
        doc = self._new_doc()

        insurer_ref = self.get("insurer_ref", "")
        matter_ref = self.get("matter_ref")
        insurer_contact = self.get("insurer_contact_name", "")
        insurer_org = self.get("insurer_contact_org", "")
        insurer_address = self.get("insurer_contact_address", "")
        insurer_email = self.get("insurer_contact_email", "")
        plaintiff = self.get("plaintiff_name")
        client = self.get("client_name")
        incident_date = self.get("incident_date", "[DATE]")
        report_date = self.get("report_date", self.get("matter_date",
                                datetime.now().strftime("%d %B %Y")))
        partner = self.get("partner_name", "the Partner")
        handler = self.get("handler_name")

        # Ref banner
        ref_tbl = doc.add_table(rows=1, cols=2)
        self._remove_table_borders(ref_tbl)
        self._cell_text(ref_tbl.rows[0].cells[0], "PRIVILEGED & CONFIDENTIAL",
                        bold=True, size=10)
        right_ref = "\n".join(filter(None, [
            f"Your ref: {insurer_ref}" if insurer_ref and insurer_ref != "[TBC]" else "",
            f"Our ref: {matter_ref}",
        ]))
        self._cell_text(ref_tbl.rows[0].cells[1], right_ref, size=10,
                        align=WD_ALIGN_PARAGRAPH.RIGHT)

        self._para(doc, "")

        # Addressee
        addr = "\n".join(filter(None, [insurer_contact, insurer_org, insurer_address]))
        if addr:
            self._para(doc, addr, size=11, sp_after=3)
        if insurer_email and insurer_email != "[TBC]":
            self._para(doc, f"BY EMAIL: {insurer_email}", bold=True, size=11, sp_after=6)

        self._para(doc, "")
        self._para(doc, "PRELIMINARY REPORT", bold=True, size=13,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, report_date, size=11, sp_after=6)
        self._para(doc, "")

        # Matter details table
        details = [
            ("Company Ref:", self.get("company_ref", "TBC")),
            ("Insured:", client),
            ("Plaintiff:", plaintiff),
            ("Policy No:", self.get("policy_no", "[POLICY NO]")),
            ("Policy Type:", self.get("policy_type", "[EL / PL / RTA / MED NEG]")),
            ("Date of Incident:", incident_date),
            ("Date of Notification:", self.get("date_notification", "TBC")),
            ("Date of Instruction:", self.get("date_instruction", "[DATE]")),
            ("Proceedings Issued:", self.get("proceedings_issued", "[DETAILS]")),
        ]
        det_tbl = doc.add_table(rows=len(details), cols=2)
        self._remove_table_borders(det_tbl)
        for i, (label, val) in enumerate(details):
            self._cell_text(det_tbl.rows[i].cells[0], label, bold=True, size=11)
            self._cell_text(det_tbl.rows[i].cells[1], val, size=11)

        self._para(doc, "")
        dear = insurer_contact.split()[0] if insurer_contact and insurer_contact != "[TBC]" else "Sir/Madam"
        self._para(doc, f"Dear {dear}", size=11, sp_after=6)
        self._para(doc,
            "We thank you for your valued instructions in relation to this matter and provide our "
            "preliminary report following review of papers received to date:-", size=11, sp_after=6)

        # Reserves summary table
        self._para(doc, "")
        res_data = [
            ("Claim Reserve:", f"€{self.get('reserve_claim', '[AMOUNT]')}"),
            ("Defence Costs Reserve:", f"€{self.get('reserve_defence_costs', '[AMOUNT]')} plus VAT"),
            ("Claim Paid:", self.get("claim_paid", "NIL")),
            ("DCR Costs Paid:", self.get("dcr_costs_paid", "NIL")),
            ("Claimants' Cost Reserve:", f"€{self.get('reserve_claimants_costs', '[AMOUNT]')} plus VAT"),
        ]
        res_tbl = doc.add_table(rows=len(res_data), cols=2)
        res_tbl.style = "Table Grid"
        for i, (label, val) in enumerate(res_data):
            self._cell_text(res_tbl.rows[i].cells[0], label, bold=True, size=10)
            self._cell_text(res_tbl.rows[i].cells[1], val, size=10)

        self._para(doc, "")

        # Report sections
        sections = [
            ("Executive Summary",
             self.get("executive_summary",
                      f"[INSERT EXECUTIVE SUMMARY — who the Plaintiff is, "
                      f"what happened on {incident_date}, what proceedings have issued]")),
            ("Incident Circumstances",
             self.get("incident_circumstances",
                      "[INSERT FULL DESCRIPTION OF THE INCIDENT CIRCUMSTANCES]")),
            ("PIAB",
             self.get("piab_details",
                      "[INSERT PIAB DETAILS — Form A, Form B, PIAB Auth No., date of authorisation]")),
            ("Procedural Position",
             self.get("procedural_position",
                      "[INSERT CURRENT PROCEDURAL POSITION — appearances entered, correspondence etc.]")),
            ("Liability",
             self.get("liability_section",
                      "[INSERT LIABILITY ANALYSIS — alleged breaches, insured's position, "
                      "contributory negligence, indemnity issues]")),
            ("Quantum",
             self.get("quantum_section",
                      "[INSERT QUANTUM — injuries per summons, medical reports, "
                      "guidelines bracket, loss of earnings, special damages]")),
        ]
        for heading, content in sections:
            self._heading(doc, heading, size=11, underline=True)
            self._para(doc, content, size=11, sp_after=6)

        # Reserve breakdown table
        self._para(doc, "")
        self._heading(doc, "Claim Reserve", size=11, underline=True)
        res_breakdown = [
            ("Item", "Estimated Claim Reserve"),
            ("General Damages", f"€{self.get('reserve_general_damages', '[AMOUNT]')}"),
            ("Loss of Earnings", f"€{self.get('reserve_loss_of_earnings', '[AMOUNT]')}"),
            ("Special Damages", f"€{self.get('reserve_special_damages', '[AMOUNT]')}"),
            ("ESTIMATED TOTAL CLAIM RESERVE", f"€{self.get('reserve_claim', '[AMOUNT]')}"),
        ]
        rb_tbl = doc.add_table(rows=len(res_breakdown), cols=2)
        rb_tbl.style = "Table Grid"
        for i, (item, val) in enumerate(res_breakdown):
            is_header = i == 0
            is_total = i == len(res_breakdown) - 1
            self._cell_text(rb_tbl.rows[i].cells[0], item, bold=(is_header or is_total), size=10)
            self._cell_text(rb_tbl.rows[i].cells[1], val, bold=(is_header or is_total), size=10)
            if is_header:
                self._cell_bg(rb_tbl.rows[i].cells[0], "D9D9D9")
                self._cell_bg(rb_tbl.rows[i].cells[1], "D9D9D9")

        # Defence Cost Reserve
        self._para(doc, "")
        self._heading(doc, "Defence Cost Reserve", size=11, underline=True)
        self._para(doc,
            f"We recommend reserving a professional fee of €{self.get('reserve_defence_costs', '[AMOUNT]')} "
            f"plus VAT in respect of the Defence Costs.",
            size=11, sp_after=6)

        # Recommendations
        self._heading(doc, "Recommendations", size=11, underline=True)
        recs_text = self.get("recommendations",
            "[INSERT RECOMMENDATIONS — e.g. enter appearance, brief counsel, instruct engineer, "
            "obtain RBA, arrange medical examination, raise Notice for Particulars]")
        for rec in recs_text.split("\n"):
            if rec.strip():
                p = doc.add_paragraph(style="List Number")
                p.paragraph_format.space_after = Pt(3)
                r = p.add_run(rec.strip())
                r.font.name = "Calibri"
                r.font.size = Pt(11)

        self._para(doc, "")
        self._para(doc,
            f"We look forward to receiving instructions and in the interim if there are any queries, "
            f"please contact {partner} or {handler} of this office.", size=11, sp_after=6)
        self._para(doc, "Kind regards", size=11)
        self._para(doc, "")
        self._para(doc, "Yours sincerely", size=11)
        self._para(doc, "")
        self._para(doc, "")
        self._para(doc, "DWF (Ireland) LLP", bold=True, size=11)

        ref = self.get("matter_ref", "report")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Preliminary_Report"))

    # ─────────────────────────────────────────────────────────────
    # 4. BRIEF COVER PAGE
    # ─────────────────────────────────────────────────────────────

    def gen_brief_cover(self):
        doc = self._new_doc()

        court = self.get("court_type", "High Court")
        plaintiff = self.get("plaintiff_name")
        client = self.get("client_name")
        record_no = self.get("record_no", "")
        adj_no = self.get("adj_no", "")
        ca_no = self.get("ca_no", "")
        matter_ref = self.get("matter_ref", "")
        brief_type = self.get("brief_type", "BRIEF TO COUNSEL")

        is_wrc = "wrc" in court.lower() or "workplace" in court.lower()

        if is_wrc:
            court_display = "WORKPLACE RELATIONS COMMISSION"
            p1_label, p2_label = "COMPLAINANT", "RESPONDENT"
        elif "circuit" in court.lower():
            court_display = "THE CIRCUIT COURT"
            p1_label, p2_label = "PLAINTIFF", "DEFENDANT"
        elif "district" in court.lower():
            court_display = "THE DISTRICT COURT"
            p1_label, p2_label = "PLAINTIFF", "DEFENDANT"
        else:
            court_display = "THE HIGH COURT"
            p1_label, p2_label = "PLAINTIFF", "DEFENDANT"

        for _ in range(4):
            self._para(doc, "")

        self._para(doc, court_display, bold=True, size=18,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12)

        if is_wrc:
            if adj_no and adj_no != "[TBC]":
                self._para(doc, f"ADJ: {adj_no}", bold=True, size=14,
                           align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=4)
            if ca_no and ca_no != "[TBC]":
                self._para(doc, f"CA: {ca_no}", bold=True, size=14,
                           align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12)
        elif record_no and record_no != "[TBC]":
            self._para(doc, f"Record No. {record_no}", bold=True, size=14,
                       align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12)

        self._para(doc, "BETWEEN/", bold=True, size=13,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=8)

        self._para(doc, plaintiff.upper(), bold=True, size=15,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, p1_label, size=12,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=8)
        self._para(doc, "-AND-", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=8)
        self._para(doc, client.upper(), bold=True, size=15,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, p2_label, size=12,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=30)

        for _ in range(3):
            self._para(doc, "")

        self._para(doc, brief_type.upper(), bold=True, size=20,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=40)

        for _ in range(4):
            self._para(doc, "")

        self._para(doc, "DWF (Ireland) LLP", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        for line in ["The Lennox", "50 Richmond Street South", "Saint Kevin's", "Dublin 2"]:
            self._para(doc, line, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        if matter_ref and matter_ref != "[TBC]":
            self._para(doc, f"(Ref: {matter_ref})", size=11,
                       align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)

        ref = self.get("matter_ref", "brief")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Brief_Cover"))

    # ─────────────────────────────────────────────────────────────
    # 5. INDEX TO BRIEF
    # ─────────────────────────────────────────────────────────────

    def gen_index_brief(self):
        doc = self._new_doc()

        court = self.get("court_type", "High Court")
        plaintiff = self.get("plaintiff_name")
        client = self.get("client_name")
        record_no = self.get("record_no", "")
        matter_ref = self.get("matter_ref", "")
        add_defs = self.get("additional_defendants", "")
        defended = self.get("defended_party", "First Named Defendant")

        # Caption
        if "circuit" in court.lower():
            court_disp = "THE CIRCUIT COURT"
        elif "district" in court.lower():
            court_disp = "THE DISTRICT COURT"
        else:
            court_disp = "THE HIGH COURT"

        self._para(doc, court_disp, bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        if record_no and record_no != "[TBC]":
            self._para(doc, f"Record No {record_no}", bold=True, size=13,
                       align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=6)
        self._para(doc, "BETWEEN:", bold=True, size=13,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, plaintiff.upper(), bold=True, size=13,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, "PLAINTIFF", size=11,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, "-AND-", bold=True, size=13,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, client.upper(), bold=True, size=13,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        if add_defs and add_defs != "[TBC]":
            self._para(doc, add_defs.upper(), bold=True, size=12,
                       align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, "DEFENDANT(S)", size=11,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=6)

        sep = doc.add_paragraph("_" * 80)
        sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sep.paragraph_format.space_after = Pt(6)

        self._para(doc, "INDEX TO BOOK", bold=True, size=16,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, "OF PAPERS", bold=True, size=16,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)

        sep2 = doc.add_paragraph("_" * 80)
        sep2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sep2.paragraph_format.space_after = Pt(12)

        self._para(doc, "")

        # Index table
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        for i, h in enumerate(["", "Date", "Page No."]):
            self._cell_text(tbl.rows[0].cells[i], h, bold=True, size=10)
            self._cell_bg(tbl.rows[0].cells[i], "D9D9D9")
        tbl.columns[0].width = Cm(10)
        tbl.columns[1].width = Cm(3)
        tbl.columns[2].width = Cm(3)

        def sec_header(letter, title):
            row = tbl.add_row()
            self._cell_text(row.cells[0], f"{letter}.    {title}", bold=True, size=10)
            for c in row.cells:
                self._cell_bg(c, "F2F2F2")

        def doc_row(name, date="", pages=""):
            row = tbl.add_row()
            self._cell_text(row.cells[0], f"     {name}", size=10)
            self._cell_text(row.cells[1], date, size=10)
            self._cell_text(row.cells[2], pages, size=10)

        # Helper to parse user-entered doc lines (format: "Name | Date | Pages")
        def parse_section(key, defaults):
            raw = self.get(key, "")
            if raw and raw != "[TBC]":
                for line in raw.split("\n"):
                    line = line.strip()
                    if line:
                        parts = [p.strip() for p in line.split("|")]
                        doc_row(
                            parts[0] if len(parts) > 0 else "",
                            parts[1] if len(parts) > 1 else "",
                            parts[2] if len(parts) > 2 else "",
                        )
            else:
                for name in defaults:
                    doc_row(name)

        sec_header("A", "Pleadings")
        parse_section("section_a_docs", [
            "Personal Injuries Summons",
            "Affidavit of Verification",
            "Memorandum of Appearance",
            "Notice for Particulars",
            "Replies to Particulars",
            "Affidavit of Verification",
            "Personal Injuries Defence",
            "Affidavit of Verification",
        ])

        sec_header("B", "Documentation / Reports")
        parse_section("section_b_docs", [
            "Incident Report",
            "Witness Statements",
            "Photographs",
            "Risk Assessment / Safe Statement",
        ])

        sec_header("C", "Medical Reports / Records")
        parse_section("section_c_docs", [
            "[Medical reports to be inserted]",
        ])

        sec_header("D", "Interpartes Correspondence")
        parse_section("section_d_docs", [
            "Correspondence between DWF and Plaintiff's Solicitors",
        ])

        extra_sec = self.get("section_e_title", "")
        if extra_sec and extra_sec != "[TBC]":
            sec_header("E", extra_sec)
            parse_section("section_e_docs", [])

        self._para(doc, "")
        self._dwf_address_block(doc)

        ref = self.get("matter_ref", "brief")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Index_to_Brief"))

    # ─────────────────────────────────────────────────────────────
    # 6. STANDARD LETTER (DWF Letterhead)
    # ─────────────────────────────────────────────────────────────

    def gen_standard_letter(self):
        doc = self._new_doc()

        rec_name = self.get("recipient_name", "")
        rec_firm = self.get("recipient_firm", "")
        rec_addr = self.get("recipient_address", "")

        self._ext_letterhead_table(doc, rec_name, rec_firm, rec_addr)
        self._para(doc, "")

        salutation = self.get("salutation", "Dear Sirs")
        self._para(doc, salutation, size=11, sp_after=6)
        self._matter_title(doc)

        body = self.get("letter_body", "[INSERT LETTER BODY]")
        for para in body.split("\n\n"):
            para = para.strip()
            if para:
                self._para(doc, para, size=11, sp_after=6)

        closing = self.get("closing", "Yours sincerely,")
        self._para(doc, "")
        self._para(doc, closing, size=11)
        self._para(doc, "")
        self._para(doc, "")
        self._dwf_sig_block(doc)

        ref = self.get("matter_ref", "letter")
        letter_type = self.get("letter_type", "Letter").replace(" ", "_")
        return self._save(doc, self._safe_filename(ref, letter_type, "Letter"))

    # ─────────────────────────────────────────────────────────────
    # 7. RBA FORM
    # ─────────────────────────────────────────────────────────────

    def gen_rba_form(self):
        doc = self._new_doc()

        matter_ref = self.get("matter_ref", "")
        plaintiff = self.get("plaintiff_name", "")
        handler = self.get("handler_name")
        handler_email = self.get("handler_email")
        handler_phone = self.get("handler_phone", "01 790 9400")

        self._para(doc, "REQUEST FOR A STATEMENT OF RECOVERABLE BENEFITS", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, "Social Welfare Services Office", size=12,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc,
            "Please use a BLACK ballpoint pen and BLOCK capitals when completing this form.",
            italic=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12)

        # Section 1: Compensator
        self._heading(doc, "SECTION 1 — COMPENSATOR DETAILS (DWF as Case Manager)")
        comp_tbl = doc.add_table(rows=6, cols=2)
        comp_tbl.style = "Table Grid"
        for i, (label, val) in enumerate([
            ("Reference no.:", matter_ref),
            ("Name of compensator / insurer:", self.get("insurer_name_full",
                                                         self.get("insurer_name", "[INSURER NAME]"))),
            ("Address of insurer:", self.get("insurer_address", "[INSURER ADDRESS]")),
            ("Name of case manager:", "DWF (Ireland) LLP, The Lennox, 50 Richmond Street South, "
                                      "Saint Kevin's, Dublin 2"),
            ("DWF Reference:", matter_ref),
            ("DWF Contact:", f"{handler}  |  {handler_email}  |  {handler_phone}"),
        ]):
            self._cell_text(comp_tbl.rows[i].cells[0], label, bold=True, size=10)
            self._cell_text(comp_tbl.rows[i].cells[1], val, size=10)

        self._para(doc, "")

        # Declaration
        self._heading(doc, "DECLARATION")
        self._para(doc,
            "In accordance with Part 11B of the Social Welfare Consolidation Act 2005, I/we wish to "
            "apply for a statement of recoverable benefits in respect of the injured person named below:",
            size=11, sp_after=6)
        sig_tbl = doc.add_table(rows=2, cols=2)
        sig_tbl.style = "Table Grid"
        self._cell_text(sig_tbl.rows[0].cells[0], "Signature of requestor:", bold=True, size=10)
        self._cell_text(sig_tbl.rows[0].cells[1], "", size=10)
        self._cell_text(sig_tbl.rows[1].cells[0], "Date:", bold=True, size=10)
        self._cell_text(sig_tbl.rows[1].cells[1], "", size=10)

        self._para(doc, "")

        # Section 2: Injured Person
        self._heading(doc, "SECTION 2 — INJURED PERSON DETAILS")

        # Auto-split plaintiff name
        parts = plaintiff.split(" ", 1)
        forenames = parts[0] if parts else ""
        surname = parts[1] if len(parts) > 1 else ""

        inj_tbl = doc.add_table(rows=7, cols=2)
        inj_tbl.style = "Table Grid"
        for i, (label, val) in enumerate([
            ("* PPS Number:", self.get("plaintiff_pps", "[PPS NUMBER - MANDATORY]")),
            ("* Surname:", self.get("plaintiff_surname", surname)),
            ("* First name(s):", self.get("plaintiff_forenames", forenames)),
            ("Gender:", self.get("plaintiff_gender", "Male  /  Female  (delete as applicable)")),
            ("* Date of birth:", self.get("plaintiff_dob", "[DD/MM/YYYY - MANDATORY]")),
            ("Address:", self.get("plaintiff_address", "[PLAINTIFF ADDRESS]")),
            ("Date of death (if applicable):", "N/A"),
        ]):
            self._cell_text(inj_tbl.rows[i].cells[0], label, bold=True, size=10)
            self._cell_text(inj_tbl.rows[i].cells[1], val, size=10)

        self._para(doc, "* Mandatory fields", italic=True, size=9, sp_before=3)
        self._para(doc, "")

        # Section 3: Injury Details
        self._heading(doc, "SECTION 3 — INJURY DETAILS")
        injuries_desc = self.get("injuries_description",
            "[DESCRIBE INJURIES SPECIFICALLY — e.g. INJURIES TO LEFT SHOULDER AND NECK REGION. "
            "Do NOT use: 'soft tissue injury', 'RTA', 'medical negligence', 'post op']")
        inj_details = doc.add_table(rows=3, cols=2)
        inj_details.style = "Table Grid"
        for i, (label, val) in enumerate([
            ("* Date of incident:", self.get("incident_date", "[DD/MM/YYYY]")),
            ("Time of incident:", self.get("incident_time", "")),
            ("Brief description of personal injury:", injuries_desc),
        ]):
            self._cell_text(inj_details.rows[i].cells[0], label, bold=True, size=10)
            self._cell_text(inj_details.rows[i].cells[1], val, size=10)

        self._para(doc, "")
        self._para(doc,
            "Important: Do NOT describe injuries as 'soft tissue injury', 'post op', 'RTA' or "
            "'medical negligence'. Provide a specific anatomical description (e.g. 'INJURIES TO "
            "LEFT SHOULDER AND NECK REGION').",
            italic=True, size=9, sp_after=3)

        return self._save(doc, self._safe_filename(matter_ref, plaintiff, "RBA_Form"))

    # ─────────────────────────────────────────────────────────────
    # 8. NOTICE FOR PARTICULARS
    # ─────────────────────────────────────────────────────────────

    def gen_notice_particulars(self):
        doc = self._new_doc()
        self._court_caption(doc, font="Times New Roman")

        plaintiff = self.get("plaintiff_name")
        matter_ref = self.get("matter_ref", "")
        defended = self.get("defended_party", "Defendant")
        pl_solic_firm = self.get("plaintiff_solicitor_firm", "[Plaintiff's Solicitors]")
        pl_solic_addr = self.get("plaintiff_solicitor_address", "[Address]")
        incident_date = self.get("incident_date", "[DATE]")

        self._para(doc, "NOTICE FOR PARTICULARS", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12, font="Times New Roman")

        self._para(doc,
            f"The Defendant(s) hereby call upon the Plaintiff to furnish them with the following "
            f"particulars of the Plaintiff's claim herein within twenty-one (21) days from the "
            f"date of service of this notice:",
            size=11, sp_after=6, font="Times New Roman")

        particulars = [
            "Full name, date of birth, address and occupation of the Plaintiff at the time of the "
            "alleged incident.",
            "The precise time, date and location of the alleged incident.",
            "Full and precise particulars of how the alleged incident occurred.",
            "Full particulars of the alleged breach(es) of duty (including breach of statutory duty) "
            "alleged against each Defendant.",
            "Full particulars of each and every injury alleged to have been sustained by the Plaintiff, "
            "specifying whether same is permanent or temporary.",
            "The name and address of every medical practitioner who has attended the Plaintiff in "
            "connection with the alleged injuries.",
            "Full particulars of the loss of earnings claimed including the Plaintiff's salary/rate "
            "of pay at the date of the alleged incident and the period of absence from work.",
            "Full particulars of all Special Damages claimed.",
            "Full particulars of any previous injuries, disabilities or pre-existing conditions from "
            "which the Plaintiff suffered prior to the alleged incident.",
            "Full particulars of any previous accidents or incidents involving the Plaintiff.",
            "Whether the Plaintiff has previously made any claim for personal injuries and if so, "
            "full particulars thereof.",
            "Whether proceedings have been instituted or are intended to be instituted against any "
            "other party in connection with the alleged incident and if so, full particulars thereof.",
        ]

        # Append any custom particulars
        custom = self.get("custom_particulars", "")
        if custom and custom != "[TBC]":
            for line in custom.split("\n"):
                line = line.strip()
                if line:
                    particulars.append(line)

        for item in particulars:
            p = doc.add_paragraph(style="List Number")
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(item)
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)

        self._para(doc, "")
        date = self.get("matter_date", datetime.now().strftime("%d %B %Y"))
        self._para(doc, f"DATED this {date}", size=11, font="Times New Roman", sp_after=12)
        self._para(doc, "")
        self._para(doc, "")
        self._dwf_address_block(doc, font="Times New Roman")
        self._para(doc, "")
        self._para(doc, "TO:", bold=True, size=11, font="Times New Roman")
        self._para(doc, pl_solic_firm, size=11, font="Times New Roman", sp_after=0)
        self._para(doc, pl_solic_addr, size=11, font="Times New Roman", sp_after=3)
        self._para(doc, "Solicitors for the Plaintiff", italic=True, size=11, font="Times New Roman")

        ref = self.get("matter_ref", "nfp")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Notice_for_Particulars"))

    # ─────────────────────────────────────────────────────────────
    # 9 & 10. MEMORANDUM OF APPEARANCE
    # ─────────────────────────────────────────────────────────────

    def gen_appearance_hc(self):
        return self._gen_appearance("High Court")

    def gen_appearance_cc(self):
        return self._gen_appearance("Circuit Court")

    def _gen_appearance(self, court_override):
        doc = self._new_doc()

        # Temporarily override court_type for caption
        orig = self.data.get("court_type", "")
        self.data["court_type"] = court_override
        self._court_caption(doc, font="Times New Roman")
        self.data["court_type"] = orig

        plaintiff = self.get("plaintiff_name")
        client = self.get("client_name")
        matter_ref = self.get("matter_ref", "")
        defended = self.get("defended_party", "Defendant")

        self._para(doc, "MEMORANDUM OF APPEARANCE", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12, font="Times New Roman")

        self._para(doc,
            f"ENTER an Appearance for the {defended}, {client}, in the above matter.",
            size=11, font="Times New Roman", sp_after=12)

        date = self.get("matter_date", datetime.now().strftime("%d %B %Y"))
        self._para(doc, f"DATED this {date}", size=11, font="Times New Roman", sp_after=16)
        self._para(doc, "")
        self._para(doc, "")
        self._dwf_address_block(doc, font="Times New Roman")

        court_short = "HC" if "high" in court_override.lower() else "CC"
        ref = self.get("matter_ref", "app")
        return self._save(doc, self._safe_filename(ref, plaintiff, f"Appearance_{court_short}"))

    # ─────────────────────────────────────────────────────────────
    # 11. NOTICE OF MOTION
    # ─────────────────────────────────────────────────────────────

    def gen_motion(self):
        doc = self._new_doc()
        self._court_caption(doc, font="Times New Roman")

        plaintiff   = self.get("plaintiff_name")
        client      = self.get("client_name")
        matter_ref  = self.get("matter_ref", "")
        return_date = self.get("lit_return_date", "[RETURN DATE]")
        grounds_raw = self.get("lit_grounds", "[GROUNDS TO BE INSERTED]")
        relief_raw  = self.get("lit_relief_sought", "[RELIEF SOUGHT TO BE INSERTED]")
        pl_firm     = self.get("plaintiff_solicitor_firm", "[Plaintiff's Solicitors]")
        pl_addr     = self.get("plaintiff_solicitor_address", "[Address]")

        self._para(doc, "NOTICE OF MOTION", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12, font="Times New Roman")

        self._para(doc,
            f"TAKE NOTICE that on {return_date} or so soon thereafter as Counsel may be heard, "
            f"the Defendant/Applicant, {client}, will apply to this Honourable Court for the "
            f"following relief:",
            size=11, sp_after=6, font="Times New Roman")

        self._heading(doc, "RELIEF SOUGHT")
        for item in relief_raw.split("\n"):
            item = item.strip()
            if item:
                p = doc.add_paragraph(style="List Number")
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(item)
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        self._para(doc, "")
        self._heading(doc, "GROUNDS")
        self._para(doc,
            "The grounds upon which this application is made are as follows:",
            size=11, sp_after=4, font="Times New Roman")
        for item in grounds_raw.split("\n"):
            item = item.strip()
            if item:
                p = doc.add_paragraph(style="List Number")
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(item)
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        self._para(doc, "")
        date = self.get("matter_date", datetime.now().strftime("%d %B %Y"))
        self._para(doc, f"DATED this {date}", size=11, font="Times New Roman", sp_after=16)
        self._para(doc, "")
        self._para(doc, "")
        self._dwf_address_block(doc, font="Times New Roman")
        self._para(doc, "")
        self._para(doc, "TO:", bold=True, size=11, font="Times New Roman")
        self._para(doc, pl_firm, size=11, font="Times New Roman", sp_after=0)
        self._para(doc, pl_addr, size=11, font="Times New Roman", sp_after=3)
        self._para(doc, "Solicitors for the Plaintiff", italic=True, size=11, font="Times New Roman")

        ref = self.get("matter_ref", "motion")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Notice_of_Motion"))

    # ─────────────────────────────────────────────────────────────
    # 12. NOTICE OF OPPOSITION
    # ─────────────────────────────────────────────────────────────

    def gen_opposition(self):
        doc = self._new_doc()
        self._court_caption(doc, font="Times New Roman")

        plaintiff   = self.get("plaintiff_name")
        client      = self.get("client_name")
        matter_ref  = self.get("matter_ref", "")
        return_date = self.get("lit_return_date", "[RETURN DATE]")
        grounds_raw = self.get("lit_opposition_grounds",
                               self.get("lit_grounds", "[GROUNDS OF OPPOSITION TO BE INSERTED]"))
        pl_firm = self.get("plaintiff_solicitor_firm", "[Plaintiff's Solicitors]")
        pl_addr = self.get("plaintiff_solicitor_address", "[Address]")

        self._para(doc, "NOTICE OF OPPOSITION", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12, font="Times New Roman")

        self._para(doc,
            f"TAKE NOTICE that the Plaintiff, {plaintiff}, will, on {return_date} or so soon "
            f"thereafter as Counsel may be heard, oppose the motion filed herein on the "
            f"following grounds:",
            size=11, sp_after=6, font="Times New Roman")

        for item in grounds_raw.split("\n"):
            item = item.strip()
            if item:
                p = doc.add_paragraph(style="List Number")
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(item)
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        self._para(doc, "")
        date = self.get("matter_date", datetime.now().strftime("%d %B %Y"))
        self._para(doc, f"DATED this {date}", size=11, font="Times New Roman", sp_after=16)
        self._para(doc, "")
        self._para(doc, "")
        self._dwf_address_block(doc, font="Times New Roman")
        self._para(doc, "")
        self._para(doc, "TO:", bold=True, size=11, font="Times New Roman")
        self._para(doc, pl_firm, size=11, font="Times New Roman", sp_after=0)
        self._para(doc, pl_addr, size=11, font="Times New Roman", sp_after=3)
        self._para(doc, "Solicitors for the Plaintiff", italic=True, size=11, font="Times New Roman")

        ref = self.get("matter_ref", "opp")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Notice_of_Opposition"))

    # ─────────────────────────────────────────────────────────────
    # 13. REPLY
    # ─────────────────────────────────────────────────────────────

    def gen_reply(self):
        doc = self._new_doc()
        self._court_caption(doc, font="Times New Roman")

        plaintiff  = self.get("plaintiff_name")
        client     = self.get("client_name")
        matter_ref = self.get("matter_ref", "")
        reply_raw  = self.get("lit_reply_grounds",
                              self.get("lit_grounds", "[REPLY CONTENT TO BE INSERTED]"))
        pl_firm    = self.get("plaintiff_solicitor_firm", "[Plaintiff's Solicitors]")
        pl_addr    = self.get("plaintiff_solicitor_address", "[Address]")

        self._para(doc, "REPLY", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12, font="Times New Roman")

        self._para(doc,
            f"The Defendant/Respondent, {client}, by its Solicitors, DWF (Ireland) LLP, "
            f"replies to the Plaintiff's Opposition as follows:",
            size=11, sp_after=6, font="Times New Roman")

        for item in reply_raw.split("\n"):
            item = item.strip()
            if item:
                p = doc.add_paragraph(style="List Number")
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(item)
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        self._para(doc, "")
        date = self.get("matter_date", datetime.now().strftime("%d %B %Y"))
        self._para(doc, f"DATED this {date}", size=11, font="Times New Roman", sp_after=16)
        self._para(doc, "")
        self._para(doc, "")
        self._dwf_address_block(doc, font="Times New Roman")
        self._para(doc, "")
        self._para(doc, "TO:", bold=True, size=11, font="Times New Roman")
        self._para(doc, pl_firm, size=11, font="Times New Roman", sp_after=0)
        self._para(doc, pl_addr, size=11, font="Times New Roman", sp_after=3)
        self._para(doc, "Solicitors for the Plaintiff", italic=True, size=11, font="Times New Roman")

        ref = self.get("matter_ref", "reply")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Reply"))

    # ─────────────────────────────────────────────────────────────
    # 14. MEDIATION BRIEF / POSITION PAPER
    # ─────────────────────────────────────────────────────────────

    def gen_mediation_brief(self):
        doc = self._new_doc()

        plaintiff      = self.get("plaintiff_name")
        client         = self.get("client_name")
        matter_ref     = self.get("matter_ref", "")
        mediation_date = self.get("lit_mediation_date", "[DATE OF MEDIATION]")
        mediator       = self.get("lit_mediator_name", "[MEDIATOR NAME]")
        incident_date  = self.get("incident_date", "[DATE]")
        record_no      = self.get("record_no", "")
        date           = self.get("matter_date", datetime.now().strftime("%d %B %Y"))

        self._para(doc, "PRIVILEGED AND WITHOUT PREJUDICE", bold=True, size=11,
                   align=WD_ALIGN_PARAGRAPH.CENTER, color=(214, 0, 86), sp_after=3)
        self._para(doc, "MEDIATION POSITION PAPER", bold=True, size=16,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, f"{plaintiff.upper()} -v- {client.upper()}", bold=True, size=12,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        if record_no and record_no != "[TBC]":
            self._para(doc, f"Record No. {record_no}", size=11,
                       align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, f"Mediation Date: {mediation_date}", size=11,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=3)
        self._para(doc, f"Mediator: {mediator}", size=11,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12)

        det_tbl = doc.add_table(rows=4, cols=2)
        self._remove_table_borders(det_tbl)
        det_tbl.columns[0].width = Cm(5)
        det_tbl.columns[1].width = Cm(11)
        for i, (lbl, val) in enumerate([
            ("Prepared by:", "DWF (Ireland) LLP"),
            ("On behalf of:", client),
            ("Date:", date),
            ("Our Ref:", matter_ref),
        ]):
            self._cell_text(det_tbl.rows[i].cells[0], lbl, bold=True, size=10)
            self._cell_text(det_tbl.rows[i].cells[1], val, size=10)
        self._para(doc, "")

        self._heading(doc, "1. BACKGROUND", size=12)
        self._para(doc,
            self.get("executive_summary",
                f"This mediation concerns a personal injuries action arising from an incident "
                f"on {incident_date}. The Plaintiff, {plaintiff}, brings this claim against "
                f"the Defendant, {client}."),
            size=11, sp_after=6)

        self._heading(doc, "2. CLIENT'S POSITION", size=12)
        self._para(doc,
            self.get("lit_position_statement",
                self.get("liability_section",
                    "[INSERT CLIENT'S POSITION ON LIABILITY AND INDEMNITY]")),
            size=11, sp_after=6)

        self._heading(doc, "3. KEY ISSUES IN DISPUTE", size=12)
        for item in self.get("lit_key_issues",
                "Liability\nQuantum\nContributory Negligence\nSpecial Damages").split("\n"):
            item = item.strip()
            if item:
                p = doc.add_paragraph(style="List Bullet")
                p.paragraph_format.space_after = Pt(3)
                r = p.add_run(item)
                r.font.name = "Calibri"
                r.font.size = Pt(11)

        self._heading(doc, "4. QUANTUM", size=12)
        for para in self.get("quantum_section",
                "[INSERT QUANTUM ANALYSIS — injuries, medical evidence, PIAB award, "
                "bracket, loss of earnings, special damages]").split("\n\n"):
            if para.strip():
                self._para(doc, para.strip(), size=11, sp_after=6)

        self._heading(doc, "5. RESERVE POSITION", size=12)
        res_tbl = doc.add_table(rows=2, cols=2)
        self._remove_table_borders(res_tbl)
        for i, (lbl, val) in enumerate([
            ("Claim Reserve:", f"€{self.get('reserve_claim', '[AMOUNT]')}"),
            ("Defence Costs Reserve:", f"€{self.get('reserve_defence_costs', '[AMOUNT]')} + VAT"),
        ]):
            self._cell_text(res_tbl.rows[i].cells[0], lbl, bold=True, size=10)
            self._cell_text(res_tbl.rows[i].cells[1], val, size=10)
        self._para(doc, "")

        self._heading(doc, "6. SETTLEMENT AUTHORITY", size=12)
        self._para(doc,
            self.get("lit_settlement_range",
                "[TO BE CONFIRMED BY INSURERS PRIOR TO MEDIATION — WITHOUT PREJUDICE]"),
            size=11, italic=True, sp_after=6)

        self._heading(doc, "7. RECOMMENDATIONS", size=12)
        for item in self.get("recommendations",
                "Engage constructively in the mediation process.\n"
                "Explore possibilities for compromise on quantum.").split("\n"):
            item = item.strip()
            if item:
                self._para(doc, item, size=11, sp_after=3)

        self._para(doc, "")
        self._para(doc,
            "This position paper is prepared on a without prejudice and privileged basis "
            "for the purposes of mediation only and shall not be referred to in any court proceedings.",
            italic=True, size=9, sp_after=6)
        self._para(doc, "")
        self._dwf_sig_block(doc)

        ref = self.get("matter_ref", "med")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Mediation_Brief"))

    # ─────────────────────────────────────────────────────────────
    # 15. EX PARTE APPLICATION + GROUNDING AFFIDAVIT
    # ─────────────────────────────────────────────────────────────

    def gen_ex_parte(self):
        doc = self._new_doc()
        self._court_caption(doc, font="Times New Roman")

        plaintiff          = self.get("plaintiff_name")
        client             = self.get("client_name")
        matter_ref         = self.get("matter_ref", "")
        deponent           = self.get("lit_deponent_name", self.get("handler_name", "[DEPONENT NAME]"))
        deponent_capacity  = self.get("lit_deponent_capacity", "Solicitor")
        grounds_raw        = self.get("lit_ex_parte_grounds",
                                      self.get("lit_grounds", "[GROUNDS FOR EX PARTE RELIEF]"))
        relief_raw         = self.get("lit_relief_sought", "[RELIEF SOUGHT]")
        urgent_reasons     = self.get("lit_urgent_reasons",
                                      "[STATE REASONS WHY NOTICE CANNOT BE GIVEN TO OTHER SIDE]")
        date               = self.get("matter_date", datetime.now().strftime("%d %B %Y"))

        # ── PART A: EX PARTE NOTICE ──
        self._para(doc, "EX PARTE APPLICATION", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12, font="Times New Roman")

        self._para(doc,
            f"The Defendant/Applicant, {client}, by its Solicitors, DWF (Ireland) LLP, "
            f"applies ex parte (without notice) to the Plaintiff for the following relief:",
            size=11, sp_after=6, font="Times New Roman")

        self._heading(doc, "RELIEF SOUGHT", size=12)
        for item in relief_raw.split("\n"):
            item = item.strip()
            if item:
                p = doc.add_paragraph(style="List Number")
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(item)
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        self._para(doc, "")
        self._heading(doc, "GROUNDS", size=12)
        for item in grounds_raw.split("\n"):
            item = item.strip()
            if item:
                p = doc.add_paragraph(style="List Number")
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(item)
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

        self._para(doc, "")
        self._heading(doc, "REASON FOR EX PARTE RELIEF", size=12)
        self._para(doc, urgent_reasons, size=11, sp_after=6, font="Times New Roman")

        self._para(doc, f"DATED this {date}", size=11, font="Times New Roman", sp_after=16)
        self._para(doc, "")
        self._para(doc, "")
        self._dwf_address_block(doc, font="Times New Roman")

        # ── PART B: GROUNDING AFFIDAVIT ──
        doc.add_page_break()
        self._court_caption(doc, font="Times New Roman")

        self._para(doc, "GROUNDING AFFIDAVIT", bold=True, size=14,
                   align=WD_ALIGN_PARAGRAPH.CENTER, sp_after=12, font="Times New Roman")

        self._para(doc,
            f"I, {deponent}, {deponent_capacity}, of DWF (Ireland) LLP, The Lennox, "
            f"50 Richmond Street South, Saint Kevin's, Dublin 2, aged eighteen years and upwards, "
            f"MAKE OATH and say as follows:",
            size=11, sp_after=6, font="Times New Roman")

        aff_paras = [
            f"I am a {deponent_capacity} in the firm of DWF (Ireland) LLP, Solicitors for "
            f"the Defendant in the above-entitled proceedings and I make this Affidavit from "
            f"facts within my own knowledge save where otherwise stated and where so stated I "
            f"believe the same to be true.",
            f"I make this Affidavit in support of the within ex parte application by the "
            f"Defendant for the reliefs identified herein.",
            f"[INSERT SUBSTANTIVE AVERMENTS GROUNDING THE APPLICATION]",
            f"I crave leave to refer to the papers exhibited herein upon which I will rely.",
        ]
        for para in aff_paras:
            p = doc.add_paragraph(style="List Number")
            p.paragraph_format.space_after = Pt(6)
            r = p.add_run(para)
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)

        self._para(doc, "")
        for _ in range(3):
            self._para(doc, "")
        self._para(doc, "_" * 45, size=11, font="Times New Roman")
        self._para(doc, deponent, bold=True, size=11, font="Times New Roman")
        self._para(doc, "")
        self._para(doc,
            f"SWORN by the above-named Deponent at _________________________ "
            f"on the _____ day of _____________ 20___",
            size=11, sp_after=12, font="Times New Roman")
        self._para(doc,
            "BEFORE ME, a Commissioner for Oaths / Practising Solicitor and I know the Deponent:",
            size=11, font="Times New Roman", sp_after=12)
        self._para(doc, "_" * 45, size=11, font="Times New Roman")
        self._para(doc, "Commissioner for Oaths / Practising Solicitor",
                   size=11, font="Times New Roman")

        ref = self.get("matter_ref", "expart")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Ex_Parte_Application"))

    # ─────────────────────────────────────────────────────────────
    # 16. PROFESSIONAL FEE NOTE / BILLING
    # ─────────────────────────────────────────────────────────────

    def gen_fee_note(self):
        doc = self._new_doc()

        plaintiff     = self.get("plaintiff_name")
        client        = self.get("client_name")
        client_addr   = self.get("client_address", "")
        matter_ref    = self.get("matter_ref", "")
        bill_no       = self.get("bill_no",
                                 f"DWF/{matter_ref.replace('/', '-')}/"
                                 f"{datetime.now().strftime('%Y%m%d')}")
        bill_date     = self.get("bill_date", datetime.now().strftime("%d %B %Y"))
        try:
            vat_rate = float(self.get("vat_rate", "23").replace("%", "") or 23) / 100
        except ValueError:
            vat_rate = 0.23
        billing_period = self.get("billing_period", "[PERIOD OF SERVICES]")
        handler        = self.get("handler_name")
        partner        = self.get("partner_name")

        # ── FIRM HEADER ──
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("DWF (Ireland) LLP")
        r.bold = True
        r.font.size = Pt(18)
        r.font.name = "Calibri"
        r.font.color.rgb = RGBColor(214, 0, 86)

        self._para(doc, "The Lennox, 50 Richmond Street South, Saint Kevin's, Dublin 2",
                   size=10, sp_after=0)
        self._para(doc, "Tel: +353 1 790 9400  |  dwf.law  |  VAT Reg: IE [VAT NO]",
                   size=10, sp_after=8)
        doc.add_paragraph("─" * 100).paragraph_format.space_after = Pt(8)

        # ── BILL HEADER TABLE ──
        hdr = doc.add_table(rows=1, cols=2)
        self._remove_table_borders(hdr)
        hdr.columns[0].width = Cm(10)
        hdr.columns[1].width = Cm(6)

        lc = hdr.rows[0].cells[0]
        lc.paragraphs[0].clear()
        lp = lc.paragraphs[0]
        b = lp.add_run("To:\n")
        b.bold = True; b.font.size = Pt(10); b.font.name = "Calibri"
        b2 = lp.add_run("\n".join(filter(None, [client, client_addr])))
        b2.font.size = Pt(10); b2.font.name = "Calibri"

        rc = hdr.rows[0].cells[1]
        rc.paragraphs[0].clear()
        rp = rc.paragraphs[0]
        rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        for j, line in enumerate([
            "PROFESSIONAL FEE NOTE",
            f"Bill No: {bill_no}",
            f"Date: {bill_date}",
            f"Our Ref: {matter_ref}",
        ]):
            r2 = rp.add_run(line + ("\n" if j < 3 else ""))
            r2.font.size = Pt(10 if j > 0 else 13)
            r2.font.name = "Calibri"
            r2.bold = (j == 0)
            if j == 0:
                r2.font.color.rgb = RGBColor(214, 0, 86)

        self._para(doc, "")
        self._para(doc, f"Re: {plaintiff} -v- {client}", bold=True, size=11, sp_after=3)
        self._para(doc, f"Matter Reference: {matter_ref}", size=10, sp_after=3)
        self._para(doc, f"Period of Services: {billing_period}", size=10, sp_after=8)

        # ── TIME CHARGES TABLE ──
        self._heading(doc, "PROFESSIONAL CHARGES — TIME RECORDED", size=11)
        te_tbl = doc.add_table(rows=1, cols=6)
        te_tbl.style = "Table Grid"
        for ci, h in enumerate(["Date", "Fee Earner", "Description of Work",
                                 "Hours", "Rate (€)", "Amount (€)"]):
            self._cell_text(te_tbl.rows[0].cells[ci], h, bold=True, size=9,
                            align=WD_ALIGN_PARAGRAPH.CENTER)
            self._cell_bg(te_tbl.rows[0].cells[ci], "222E40")
            for run in te_tbl.rows[0].cells[ci].paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255, 255, 255)

        for ci, w in enumerate([Cm(2), Cm(3.2), Cm(7), Cm(1.5), Cm(2), Cm(2)]):
            for row in te_tbl.rows:
                row.cells[ci].width = w

        total_fees = 0.0
        has_entries = False
        for n in range(20):
            earner    = self.data.get(f"te_earner_{n}", "").strip()
            hours_str = self.data.get(f"te_hours_{n}", "").strip()
            rate_str  = self.data.get(f"te_rate_{n}", "").strip()
            desc      = self.data.get(f"te_desc_{n}", "").strip()
            te_date   = self.data.get(f"te_date_{n}", "").strip()
            if not earner and not hours_str:
                continue
            try:
                hours  = float(hours_str or 0)
                rate   = float(rate_str or 0)
                amount = hours * rate
            except ValueError:
                hours = rate = amount = 0.0
            total_fees += amount
            has_entries = True
            row = te_tbl.add_row()
            self._cell_text(row.cells[0], te_date, size=9)
            self._cell_text(row.cells[1], earner, size=9)
            self._cell_text(row.cells[2], desc, size=9)
            self._cell_text(row.cells[3], f"{hours:.1f}", size=9,
                            align=WD_ALIGN_PARAGRAPH.CENTER)
            self._cell_text(row.cells[4], f"€{rate:,.2f}", size=9,
                            align=WD_ALIGN_PARAGRAPH.RIGHT)
            self._cell_text(row.cells[5], f"€{amount:,.2f}", size=9,
                            align=WD_ALIGN_PARAGRAPH.RIGHT)

        if not has_entries:
            row = te_tbl.add_row()
            self._cell_text(row.cells[0], "[Date]", size=9)
            self._cell_text(row.cells[1], "[Fee Earner]", size=9)
            self._cell_text(row.cells[2], "[Description of work performed]", size=9)
            self._cell_text(row.cells[3], "—", size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
            self._cell_text(row.cells[4], "€0.00", size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)
            self._cell_text(row.cells[5], "€0.00", size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)

        sub_row = te_tbl.add_row()
        for cell in sub_row.cells:
            self._cell_bg(cell, "F0ECE7")
        self._cell_text(sub_row.cells[2], "SUB-TOTAL — Professional Charges",
                        bold=True, size=9)
        self._cell_text(sub_row.cells[5], f"€{total_fees:,.2f}", bold=True, size=9,
                        align=WD_ALIGN_PARAGRAPH.RIGHT)

        self._para(doc, "")

        # ── DISBURSEMENTS TABLE ──
        self._heading(doc, "DISBURSEMENTS", size=11)
        disb_tbl = doc.add_table(rows=1, cols=3)
        disb_tbl.style = "Table Grid"
        for ci, h in enumerate(["Date", "Description", "Amount (€)"]):
            self._cell_text(disb_tbl.rows[0].cells[ci], h, bold=True, size=9,
                            align=WD_ALIGN_PARAGRAPH.CENTER)
            self._cell_bg(disb_tbl.rows[0].cells[ci], "222E40")
            for run in disb_tbl.rows[0].cells[ci].paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255, 255, 255)

        total_disb = 0.0
        has_disb   = False
        for n in range(10):
            disb_desc   = self.data.get(f"disb_desc_{n}", "").strip()
            disb_amt_s  = self.data.get(f"disb_amount_{n}", "").strip()
            disb_date   = self.data.get(f"disb_date_{n}", "").strip()
            if not disb_desc and not disb_amt_s:
                continue
            try:
                disb_amount = float(disb_amt_s.replace("€", "").replace(",", "") or 0)
            except ValueError:
                disb_amount = 0.0
            total_disb += disb_amount
            has_disb = True
            row = disb_tbl.add_row()
            self._cell_text(row.cells[0], disb_date, size=9)
            self._cell_text(row.cells[1], disb_desc, size=9)
            self._cell_text(row.cells[2], f"€{disb_amount:,.2f}", size=9,
                            align=WD_ALIGN_PARAGRAPH.RIGHT)

        if not has_disb:
            row = disb_tbl.add_row()
            self._cell_text(row.cells[1], "NIL", size=9)
            self._cell_text(row.cells[2], "€0.00", size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)

        d_sub = disb_tbl.add_row()
        for cell in d_sub.cells:
            self._cell_bg(cell, "F0ECE7")
        self._cell_text(d_sub.cells[1], "SUB-TOTAL — Disbursements", bold=True, size=9)
        self._cell_text(d_sub.cells[2], f"€{total_disb:,.2f}", bold=True, size=9,
                        align=WD_ALIGN_PARAGRAPH.RIGHT)

        self._para(doc, "")

        # ── TOTALS ──
        vat_amount  = total_fees * vat_rate
        grand_total = total_fees + vat_amount + total_disb

        totals_tbl = doc.add_table(rows=4, cols=2)
        self._remove_table_borders(totals_tbl)
        totals_tbl.columns[0].width = Cm(12)
        totals_tbl.columns[1].width = Cm(4)
        for i, (lbl, val) in enumerate([
            ("Professional Charges:", f"€{total_fees:,.2f}"),
            (f"VAT @ {int(vat_rate * 100)}% on Professional Charges:", f"€{vat_amount:,.2f}"),
            ("Disbursements (VAT exclusive):", f"€{total_disb:,.2f}"),
            ("TOTAL NOW DUE:", f"€{grand_total:,.2f}"),
        ]):
            is_total = (i == 3)
            self._cell_text(totals_tbl.rows[i].cells[0], lbl, bold=is_total, size=10,
                            align=WD_ALIGN_PARAGRAPH.RIGHT)
            self._cell_text(totals_tbl.rows[i].cells[1], val, bold=is_total, size=10,
                            align=WD_ALIGN_PARAGRAPH.RIGHT)
            if is_total:
                for cell in totals_tbl.rows[i].cells:
                    self._cell_bg(cell, "222E40")
                    for run in cell.paragraphs[0].runs:
                        run.font.color.rgb = RGBColor(255, 255, 255)

        self._para(doc, "")
        self._para(doc,
            "Payment is due within 30 days of the date of this fee note. "
            "Please quote the bill number and matter reference when making payment.",
            size=9, italic=True, sp_after=4)
        self._para(doc,
            "DWF (Ireland) LLP is a limited liability partnership registered in Ireland. "
            "Authorised by the Law Society of Ireland.",
            size=8, italic=True, sp_after=8)
        self._para(doc, "Authorised by:", size=10, sp_after=0)
        for _ in range(3):
            self._para(doc, "")
        self._para(doc, "_" * 35, size=10, sp_after=0)
        self._para(doc, partner if (partner and partner != "[TBC]") else handler,
                   bold=True, size=10, sp_after=0)
        self._para(doc, "Partner, DWF (Ireland) LLP", size=10)

        ref = self.get("matter_ref", "bill")
        return self._save(doc, self._safe_filename(ref, plaintiff, "Fee_Note"))
