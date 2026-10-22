"""Generate a realistic FAKE residential purchase agreement PDF (plus an encrypted copy)."""
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib.styles import getSampleStyleSheet
from pypdf import PdfReader, PdfWriter

OUT = Path(__file__).parent / "samples"
FIELDS = [
    ("Buyer", "Jordan T. Example and Casey R. Example"),
    ("Seller", "Pat Q. Sample"),
    ("Property Address", "1234 Placeholder Lane, Macon, GA 31201"),
    ("Tax Parcel ID", "Q071-0042"),
    ("Purchase Price", "$289,500.00"),
    ("Earnest Money", "$3,000.00"),
    ("Earnest Money Holder", "Fictional Realty Brokerage LLC"),
    ("Seller Contribution at Closing", "$5,000.00"),
    ("Financing Type", "Conventional"),
    ("Due Diligence Period", "10 days"),
    ("Earnest Money Due", "3 days"),
    ("Financing Contingency", "21 days"),
    ("Appraisal Contingency", "17 days"),
    ("Closing", "30 days"),
    ("Possession", "At closing"),
]

def build(path: Path):
    st = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(path), pagesize=letter, title="Purchase and Sale Agreement (SAMPLE)")
    s = [Paragraph("PURCHASE AND SALE AGREEMENT", st["Title"]),
         Paragraph("<i>SAMPLE DOCUMENT - all names, addresses and figures are fictional.</i>", st["Normal"]),
         Spacer(1, 12),
         Paragraph("1. Parties and Property. The undersigned buyer(s) agree to buy and the seller(s) agree to sell "
                   "the property described below on the terms of this Agreement.", st["Normal"]), Spacer(1, 8)]
    s.append(Table([[f"{k}:", v] for k, v in FIELDS[:4]], colWidths=[170, 330]))
    s += [Spacer(1, 10), Paragraph("2. Price and Earnest Money.", st["Heading4"])]
    s.append(Table([[f"{k}:", v] for k, v in FIELDS[4:9]], colWidths=[170, 330]))
    s += [Spacer(1, 10), Paragraph("3. Deadlines. All periods are counted in calendar days from the Binding Agreement Date "
          "(the date of final acceptance). A deadline falling on a Saturday or Sunday moves to the following Monday.", st["Normal"]), Spacer(1, 6)]
    s.append(Table([[f"{k}:", v] for k, v in FIELDS[9:]], colWidths=[170, 330]))
    s += [Spacer(1, 18), Paragraph("Binding Agreement Date: ____________ (filled in at acceptance)", st["Normal"]),
          Spacer(1, 30), Paragraph("Buyer signature: ______________________ &nbsp;&nbsp; Seller signature: ______________________", st["Normal"])]
    doc.build(s)

def main():
    OUT.mkdir(exist_ok=True)
    plain = OUT / "purchase_agreement_sample.pdf"
    build(plain)
    w = PdfWriter(); [w.add_page(p) for p in PdfReader(plain).pages]
    w.encrypt("demo1234", algorithm="AES-256")
    with open(OUT / "purchase_agreement_protected.pdf", "wb") as f: w.write(f)
    print("wrote", plain, "and protected copy (password: demo1234)")

if __name__ == "__main__":
    main()
