import xlwings as xw


def replace_links(input_file, output_file, old_link, new_link):
    app = xw.App(visible=False)
    app.display_alerts = False

    try:
        wb = xw.Book(input_file)

        # 1. Alle Sheets durchgehen
        for sheet in wb.sheets:
            used_range = sheet.used_range

            # Zellwerte + Formeln ersetzen
            for cell in used_range:
                if cell.formula:
                    cell.formula = str(cell.formula).replace(old_link, new_link)

                if cell.value and isinstance(cell.value, str):
                    cell.value = cell.value.replace(old_link, new_link)

        # 2. Workbook-level Links (falls Excel sie verwaltet)
        try:
            for link in wb.api.LinkSources():
                if old_link in link:
                    wb.api.ChangeLink(Name=link, NewName=link.replace(old_link, new_link), Type=1)
        except Exception:
            # LinkSources kann None zurückgeben → ignorieren
            pass

        # 3. Speichern unter neuer Datei
        wb.save(output_file)
        wb.close()

        print(f"Fertig gespeichert: {output_file}")

    finally:
        app.quit()


if __name__ == "__main__":
    # --- HARDCODED VALUES ---
    input_excel = r"C:\data\input.xlsx"
    output_excel = r"C:\data\output.xlsx"

    old_excel_file = r"C:\data\old_file.xlsx"
    new_excel_file = r"C:\data\new_file.xlsx"

    replace_links(
        input_excel,
        output_excel,
        old_excel_file,
        new_excel_file
    )
