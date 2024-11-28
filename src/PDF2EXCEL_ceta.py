import os
import tabula
import pandas as pd
import re
import logging
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Supprimer les messages de log de PDFBox
logging.getLogger("org.apache.pdfbox").setLevel(logging.ERROR)

# Rediriger stderr vers null pour supprimer les messages indésirables
class NullWriter:
    def write(self, arg):
        pass

sys.stderr = NullWriter()

def clean_table(table):
    # Supprimer les lignes qui sont probablement des entêtes, pieds de page ou reports
    table = table.dropna(how='all')  # Supprimer les lignes où tous les éléments sont NaN
    report_pattern = re.compile(r'R\s*\.?\s*e\s*\.?\s*p\s*\.?\s*o\s*\.?\s*r\s*\.?\s*t', re.IGNORECASE)
    table = table[~table.apply(lambda row: row.astype(str).str.contains(report_pattern, regex=True).any(), axis=1)]
    
    # Réinitialiser l'index après le nettoyage
    table.reset_index(drop=True, inplace=True)
    
    return table

def convert_to_float(value):
    try:
        if isinstance(value, str):
            # Remplacer les points utilisés comme séparateurs de milliers par des chaînes vides
            value = value.replace('.', '').replace(',', '.')
            return float(value)
        return float(value)
    except ValueError:
        return value

def process_pdf(pdf_path):
    # Lire le fichier PDF
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True, encoding='latin-1')

    if not tables:
        print(f"Aucune table trouvée dans le fichier {pdf_path}")
        return None

    # Concaténer toutes les tables en une seule DataFrame
    combined_table = pd.concat(tables, ignore_index=True)

    # Nettoyer la table combinée
    cleaned_table = clean_table(combined_table)

    # Supprimer les lignes d'en-tête dupliquées et les lignes de totaux dupliquées
    header_row = cleaned_table.iloc[0]
    total_row = cleaned_table.iloc[-1]
    
    cleaned_table = cleaned_table.drop_duplicates(subset=header_row.index, keep='first')
    cleaned_table = cleaned_table.drop_duplicates(subset=total_row.index, keep='last')
    
    # Convertir les cellules contenant des chiffres en valeurs numériques (float)
    for col in cleaned_table.columns:
        cleaned_table[col] = cleaned_table[col].apply(convert_to_float)
    
    return cleaned_table

def convert_pdfs(pdf_paths, treeview):
    for pdf_path in pdf_paths:
        try:
            treeview.item(pdf_path, values=(os.path.basename(pdf_path), pdf_path, "🔄"))
            treeview.update_idletasks()
            
            excel_path = pdf_path.replace('.pdf', '.xlsx')
            
            # Traiter le fichier PDF et obtenir la table nettoyée
            final_table = process_pdf(pdf_path)
            
            if final_table is None:
                raise Exception("Aucune table trouvée")
            
            # Sauvegarder la table finale dans un fichier Excel avec le bon formatage
            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                final_table.to_excel(writer, index=False, sheet_name='Sheet1')
                
                workbook  = writer.book
                worksheet = writer.sheets['Sheet1']
                
                for column in final_table:
                    column_length = max(final_table[column].astype(str).map(len).max(), len(str(column)))
                    col_idx = final_table.columns.get_loc(column)
                    worksheet.set_column(col_idx, col_idx, column_length)
                
                # Convertir les cellules de texte en nombres dans Excel
                for row_num in range(1, len(final_table) + 1):
                    for col_num in range(len(final_table.columns)):
                        cell_value = final_table.iat[row_num - 1, col_num]
                        try:
                            # Vérifier si la valeur peut être convertie en nombre
                            numeric_value = float(cell_value.replace('.', '').replace(',', '.')) if isinstance(cell_value, str) else float(cell_value)
                            if pd.notna(numeric_value) and not pd.isinf(n_value):
                                # Appliquer la nouvelle valeur de la cellule
                                worksheet.write_number(row_num, col_num, numeric_value)
                        except (ValueError, AttributeError):
                            pass
            
            treeview.item(pdf_path, values=(os.path.basename(pdf_path), pdf_path, "✔️"))
        except Exception as e:
            treeview.item(pdf_path, values=(os.path.basename(pdf_path), pdf_path, "❌"))

def browse_files(treeview):
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    for file_path in file_paths:
        treeview.insert("", tk.END, iid=file_path, values=(os.path.basename(file_path), file_path, ""))

def start_conversion(treeview):
    pdf_paths = treeview.get_children()
    
    if not pdf_paths:
        messagebox.showwarning("Avertissement", "Veuillez sélectionner au moins un fichier PDF.")
        return
    
    convert_pdfs(pdf_paths, treeview)

def remove_item(treeview, item):
    treeview.delete(item)

# Interface graphique avec Tkinter
root = tk.Tk()
root.title("PDF to Excel Converter")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

treeview = ttk.Treeview(frame, columns=("Nom du fichier", "Chemin du fichier", "État"), show="headings")
treeview.heading("Nom du fichier", text="Nom du fichier")
treeview.heading("Chemin du fichier", text="Chemin du fichier")
treeview.heading("État", text="État")
treeview.pack(pady=5)

def on_treeview_click(event):
    item = treeview.identify_row(event.y)
    if treeview.identify_column(event.x) == "#1":  # Vérifie si la colonne cliquée est la première
        remove_item(treeview, item)

treeview.bind("<Button-1>", on_treeview_click)

browse_button = tk.Button(frame, text="Parcourir", command=lambda: browse_files(treeview))
browse_button.pack(pady=5)

convert_button = tk.Button(frame, text="Convertir", command=lambda: start_conversion(treeview))
convert_button.pack(pady=5)

root.mainloop()