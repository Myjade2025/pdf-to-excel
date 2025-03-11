import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os

def extract_invoice_data(pdf_path):
    """Extrait les données essentielles d'une facture PDF."""
    invoice_data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                # Extraction des informations clés avec regex
                date_match = re.search(r'\d{2}/\d{2}/\d{4}', text)
                invoice_match = re.search(r'Facture N°\s*(\d+)', text)
                ttc_match = re.search(r'Total TTC\s*:?\s*(\d+[.,]\d+)', text)
                ht_match = re.search(r'Total HT\s*:?\s*(\d+[.,]\d+)', text)
                tva_match = re.search(r'TVA\s*:?\s*(\d+[.,]\d+)', text)
                
                # Stockage des résultats extraits
                date = date_match.group(0) if date_match else "Non trouvé"
                invoice_number = invoice_match.group(1) if invoice_match else "Non trouvé"
                ttc = ttc_match.group(1) if ttc_match else "0"
                ht = ht_match.group(1) if ht_match else "0"
                tva = tva_match.group(1) if tva_match else "0"
                
                invoice_data.append([date, "Journal", invoice_number, "Libellé Compte", ttc, ht, tva])
    
    return invoice_data

def process_pdf_folder(folder_path):
    """Traite tous les fichiers PDF d'un dossier et génère un fichier Excel."""
    all_data = []
    
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            invoice_data = extract_invoice_data(pdf_path)
            all_data.extend(invoice_data)
    
    if all_data:
        df = pd.DataFrame(all_data, columns=["Date", "Journal", "Numéro Facture", "Libellé Compte", "TTC", "HT", "TVA"])
        output_path = os.path.join(folder_path, "Factures_Excel.xlsx")
        df.to_excel(output_path, index=False)
        return output_path, df
    else:
        return None, None

def main():
    st.title("Extraction Automatique de Factures PDF vers Excel")
    st.write("Déposez vos factures dans le dossier `C:\Factures` et cliquez sur 'Scanner'.")
    
    folder_path = "C:\Factures"
    
    if st.button("Scanner le dossier et générer Excel"):
        if os.path.exists(folder_path):
            output_file, df = process_pdf_folder(folder_path)
            if output_file:
                st.write("### Données extraites :")
                st.dataframe(df)
                st.success(f"Fichier généré avec succès : {output_file}")
            else:
                st.error("Aucune facture valide trouvée dans le dossier.")
        else:
            st.error("Le dossier 'C:\Factures' n'existe pas. Veuillez le créer et y déposer des fichiers PDF.")

if __name__ == "__main__":
    main()
