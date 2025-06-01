import pandas as pd
from fpdf import FPDF
from num2words import num2words

class RentReceipt(FPDF):
    def header(self):
        self.set_font("Helvetica", size=10)

    def add_receipt(self, data):
        self.add_page()
        self.set_font("Helvetica", size=10)

        # Receipt layout (landscape on 148 x 97 mm paper)
        self.text(10, 15, f"N° {data['numero']}")  # Receipt number
        self.text(130, 15, f"#{data['montant']:.2f}")  # Amount top right
        self.text(47, 26, data['nom'])  # Client name
        self.text(40, 35, data['montant_str'])  # Amount in words
        self.text(37, 44, data['adresse'])  # Address
        self.text(28, 67, f"{data['montant']:.2f}")  # Loyer field
        self.text(28, 90, f"{data['montant']:.2f}")  # Total field
        self.text(85, 58, data['debut'])  # Start date
        self.text(125, 58, data['fin'])  # End date
        self.text(45, 93, data['ville'])  # City
        self.text(95, 93, data['date'])  # Date


# === Load Excel ===
df = pd.read_excel("rent.xlsx", sheet_name="Sheet1")  # ← Change to your actual file name

# === Prepare PDF ===
pdf = RentReceipt(orientation="L", unit="mm", format=(97, 148))

# === Loop over each row to generate receipt ===
for i, row in df.iterrows():
    receipt_data = {
        "numero": str(row["no"]),
        "nom": row["LOCATAIRE"],
        "montant": row["LOYER"],
        "montant_str": num2words(row["LOYER"], lang="ma"),
        "adresse": row["ADDRESS"],
        "debut": pd.to_datetime(row["DATE1"]).strftime("%d.%m.%Y"),
        "fin": pd.to_datetime(row["DATE2"], dayfirst=True).strftime("%d.%m.%Y"),
        "ville": row["VILLE"],
        "date": pd.to_datetime(row["DATE1"]).strftime("%d.%m.%Y"),
    }

    pdf.add_receipt(receipt_data)

# === Save final PDF ===
    # Extract month/year from first receipt's DATE
first_date = pd.to_datetime(df.iloc[0]["DATE1"], dayfirst=True)
mois = first_date.strftime("%m").lower()  # e.g., "juin"
annee = first_date.strftime("%Y")

    # Save with custom name
pdf.output(f"recu_{mois}_{annee}.pdf")

