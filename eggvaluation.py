import pandas as pd
from tkinter import Tk, filedialog, messagebox


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


def show_message(title, message):
    """Show a message box with a given title and message."""
    root = Tk()
    root.withdraw()  # Hide the main window
    messagebox.showinfo(title, message)
    root.destroy()


# Dialoge für Dateiauswahl
try:
    mastr_file = select_file("Bitte die Datei mit dem Auszug aus dem Marktstammdatenregister auswählen.")
    eeg_file = select_file("Bitte die Datei mit den Zuschlägen der EEG-Biomasseausschreibung auswählen.")
    output_file = save_file("Bitte den Speicherort für die Ausgabedatei angeben.")

    # Lesen der Excel-Dateien
    mastr_df = pd.read_excel(mastr_file)
    eeg_sheet2_df = pd.read_excel(eeg_file, sheet_name=1)
    eeg_sheet3_df = pd.read_excel(eeg_file, sheet_name=2)


    # Verarbeitung der EEG-Zuschlagsliste
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
                zuschlagsnummer = eeg_sheet3_df.loc[
                    eeg_sheet3_df['Nummern'].apply(lambda x: num in x), 'Zuschlagsnummer'].values
                return zuschlagsnummer[0] if len(zuschlagsnummer) > 0 else None
        return None


    mastr_df['Zuschlagsnummer'] = mastr_df.apply(find_matching_zuschlagsnummer, axis=1)
    matching_rows = mastr_df.dropna(subset=['Zuschlagsnummer'])

    # Spalten neu anordnen und unnötige Spalten entfernen
    columns_order = ['Zuschlagsnummer', 'MaStR-Nr. der EEG-Anlage', 'MaStR-Nr. der Einheit'] + \
                    [col for col in matching_rows.columns if col not in [
                        'Zuschlagsnummer', 'MaStR-Nr. der EEG-Anlage', 'MaStR-Nr. der Einheit', 'Gemarkung',
                        'Flurstück',
                        'Gemeindeschlüssel', 'Anzahl der Solar-Module', 'Hauptausrichtung der Solar-Module',
                        'Name des Windparks', 'Nabenhöhe der Windenergieanlage',
                        'Rotordurchmesser der Windenergieanlage',
                        'Hersteller der Windenergieanlage', 'Typenbezeichnung']]

    final_df = matching_rows[columns_order]

    # Step 6: Speichern der Ausgabe in einer neuen Excel-Datei
    final_df.to_excel(output_file, index=False)

    # Step 7: Mitteilung bei Erfolg
    show_message("Erfolg", "Die Auswertung wurde erfolgreich durchgeführt und die Datei wurde gespeichert.")

except Exception as e:
    # Step 8: Mitteilung bei Fehler
    show_message("Fehler", f"Ein Fehler ist aufgetreten: {e}")
