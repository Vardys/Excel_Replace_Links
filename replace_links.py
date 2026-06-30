import xlwings as xw


def replace_links(input_file, output_file, old_link, new_link):
    app = xw.App(visible=False)
    app.display_alerts = False

    try:
        wb = xw.Book(input_file)

        for sheet in wb.sheets:
            rng = sheet.used_range

            for cell in rng:

                # --- FORMELN (wichtig: niemals .value anfassen bei Formeln) ---
                if cell.formula:
                    try:
                        formula = cell.formula

                        if isinstance(formula, str) and old_link in formula:
                            cell.formula = formula.replace(old_link, new_link)

                    except Exception:
                        pass

                # --- NORMALE TEXTE ---
                if cell.value and isinstance(cell.value, str):
                    # nur ersetzen wenn es KEINE Formel ist
                    if not (isinstance(cell.formula, str) and cell.formula.startswith("=")):
                        cell.value = cell.value.replace(old_link, new_link)

        # --- Workbook-Level Links ---
        try:
            links = wb.api.LinkSources()
            if links:
                for link in links:
                    if old_link in link:
                        wb.api.ChangeLink(
                            Name=link,
                            NewName=link.replace(old_link, new_link),
                            Type=1
                        )
        except Exception:
            pass

        wb.save(output_file)
        wb.close()

        print(f"Fertig gespeichert: {output_file}")

    finally:
        app.quit()


if __name__ == "__main__":
    # --- HARDCODED ---
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
