import pandas as pd
from tkinter import Tk, filedialog

def select_file(prompt):
    """Open a file selection dialog and return the selected file path."""
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title=prompt)
    root.destroy()
    return file_path

def save_file(prompt):
    """Open a file save dialog and return the selected file path."""
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.asksaveasfilename(title=prompt, defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    root.destroy()
    return file_path

# Dialoge zur Dateiauswahl
mastr_file = select_file("Bitte die Datei mit dem Auszug aus dem Marktstammdatenregister auswählen.")
eeg_file = select_file("Bitte die Datei mit den Zuschlägen der EEG-Biomasseausschreibung auswählen.")
output_file = save_file("Bitte den Speicherort für die Ausgabedatei auswählen.")

# Lesen der Excel-Dateien
mastr_df = pd.read_excel(mastr_file)
eeg_sheet2_df = pd.read_excel(eeg_file, sheet_name=1)
eeg_sheet3_df = pd.read_excel(eeg_file, sheet_name=2)

# Verarbeitung von EEG-Zuschlagsliste zur Extraktion relevanter Zahlen
def extract_numbers(cell):
    if pd.isna(cell):
        return []
    numbers = cell.replace(", ", "; ").split("; ")
    return [num for num in numbers if num.startswith("EEG") or num.startswith("SEE")]

eeg_sheet3_df['Nummern'] = eeg_sheet3_df.iloc[:, 1].apply(extract_numbers)
eeg_sheet3_df['Zuschlagsnummer'] = eeg_sheet3_df.iloc[:, 0]  # Extract Zuschlagsnummer
eeg_numbers = set(num for sublist in eeg_sheet3_df['Nummern'] for num in sublist)

# Übereinstimmende Zeilen in MaStR-Datei suchen und Zuschlagsnummer hinzufügen
def find_matching_zuschlagsnummer(row):
    mastr_numbers = set([row['MaStR-Nr. der EEG-Anlage'], row['MaStR-Nr. der Einheit']])
    for num in mastr_numbers:
        if num in eeg_numbers:
            zuschlagsnummer = eeg_sheet3_df.loc[eeg_sheet3_df['Nummern'].apply(lambda x: num in x), 'Zuschlagsnummer'].values
            return zuschlagsnummer[0] if len(zuschlagsnummer) > 0 else None
    return None

mastr_df['Zuschlagsnummer'] = mastr_df.apply(find_matching_zuschlagsnummer, axis=1)
matching_rows = mastr_df.dropna(subset=['Zuschlagsnummer'])

# Speichern der Ausgabe in einer neuen Excel-Datei
matching_rows.to_excel(output_file, index=False)
