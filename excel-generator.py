import xlsxwriter
# Define the column mappings for each Excel file
excel_columns = {
    "Excel 1: Standard Naming": [
        "Référence", "Montant Total", "Date de Création", "Nom du Tireur", "Nom du Tiré",
        "Devise", "Lieu de Paiement", "Adresse du Tiré", "Garantie", "Code Barre"
    ],
    "Excel 2: Detailed Naming": [
        "Référence du Document", "Valeur Totale", "Date d'Émission", "Ville d'Émission",
        "Nom du Bénéficiaire", "RIB Bancaire du Tiré", "Devise Utilisée", "Code Réservé",
        "Adresse Complète du Tiré", "Montant en Lettres"
    ],
    "Excel 3: Business-Oriented Naming": [
        "Code Réf", "Montant (USD)", "Nom de l'Endosseur", "Nom du Bénéficiaire",
        "Date de Création", "Ville d'Origine", "Garanties", "Lieu de Paiement", "Barre Code"
    ],
    "Excel 4: Abbreviated Naming": [
        "Réf.", "Valeur ($)", "Date", "Ville", "Devise", "Tiré", "RIB", "Montant en Mots", "Endosseur"
    ],
    "Excel 5: Technical Naming": [
        "Identifiant de Référence", "Montant (USD)", "Code Réservé", "Date de Création",
        "Adresse du Tiré", "Banque (RIB)", "Lieu de Paiement", "Nom du Garant", "Code-Barre"
    ]
}

# genearte the Excel files from this 
for excel_name, columns in excel_columns.items():
    # Create a new Excel file
    workbook = xlsxwriter.Workbook(f"{excel_name}.xlsx")
    worksheet = workbook.add_worksheet()

    # Write the column headers
    for i, column in enumerate(columns):
        worksheet.write(0, i, column)

    # Close the workbook
    workbook.close()