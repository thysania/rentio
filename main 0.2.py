import pandas as pd
from fpdf import FPDF
from num2words import num2words

class RentReceipt(FPDF):
    def header(self):
        self.set_font("Helvetica", size=10)

    def add_receipt(self, data):
        self.add_page()
        self.set_font("Helvetica", size=10)

        # Receipt layout (landscape on 147 x 93 mm paper)
        self.set_font("Helvetica", size=10)
        self.text(75, 37, data['numero'])  # Receipt number
        self.text(175, 37, f"#{data['montant']:.2f}")  # Amount top right
        self.text(105, 45, data['nom'])  # Client name
        self.text(91, 52, data['montant_str'])  # Amount in words
        self.set_font("Helvetica", size=8)
        self.text(106, 69, data['adresse'])  # Address
        self.set_font("Helvetica", size=9)
        self.text(90, 77, f"{data['montant']:.2f}")  # Loyer field
        self.text(90, 116, f"{data['montant']:.2f}")  # Total field
        self.text(150, 81, data['debut'])  # Start date
        self.text(150, 87, data['fin'])  # End date
        self.text(120, 117, data['ville'])  # City
        self.text(170, 117, data['date'])  # Date


# === Load Excel ===
df = pd.read_excel("rent.xlsx", sheet_name="main")

# === Prepare PDF ===
pdf = RentReceipt(orientation="L", unit="mm", format="A5")
pdf.set_margin(0)

# === Loop over each row to generate receipt ===
for i, row in df.iterrows():
    receipt_data = {
        "numero": str(row["NO"]),
        "nom": row["LOCATAIRE"],
        "montant": row["LOYER"],
        "montant_str": num2words(row["LOYER"], lang='ma', to='cardinal').capitalize(),
        "adresse": row["ADDRESS"],
        "debut": pd.to_datetime(row["DATE1"], dayfirst=True).strftime("%d.%m.%Y"),
        "fin": pd.to_datetime(row["DATE2"], dayfirst=True).strftime("%d.%m.%Y"),
        "ville": row["VILLE"],
        "date": pd.to_datetime(row["DATE1"], dayfirst=True).strftime("%d.%m.%Y"),
    }

    pdf.add_receipt(receipt_data)

# === Save final PDF ===
    # Extract month/year
first_date = pd.to_datetime(df.iloc[0]["DATE1"], dayfirst=True)
mois = first_date.strftime("%m").lower()  # e.g., "juin"
annee = first_date.strftime("%Y")

    # Save with custom name
pdf.output(f"recu_{mois}_{annee}.pdf")