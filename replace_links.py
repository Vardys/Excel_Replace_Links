import xlwings as xw


def replace_links(input_file, output_file, old_link, new_link):
    app = xw.App(visible=False)
    app.display_alerts = False

    try:
        wb = xw.Book(input_file)

        for sheet in wb.sheets:
            used_range = sheet.api.UsedRange

            # Ganze Range als Array holen (viel stabiler als cell-by-cell)
            formulas = used_range.Formula
            values = used_range.Value

            # Excel gibt bei 1 Zelle manchmal kein 2D-Array zurück
            if not isinstance(formulas, (list, tuple)):
                formulas = [[formulas]]
                values = [[values]]

            rows = len(formulas)
            cols = len(formulas[0])

            for i in range(rows):
                for j in range(cols):

                    # 1. Formel behalten!
                    if formulas[i][j] and isinstance(formulas[i][j], str) and formulas[i][j].startswith("="):
                        formulas[i][j] = formulas[i][j].replace(old_link, new_link)

                    # 2. Normale Texte (keine Formeln!)
                    elif values[i][j] and isinstance(values[i][j], str):
                        values[i][j] = values[i][j].replace(old_link, new_link)

            # Zurückschreiben OHNE Formeln zu zerstören
            used_range.Formula = formulas
            used_range.Value = values

        # Workbook Links (Excel interne Verknüpfungen)
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
    input_excel = r"C:\data\input.xlsx"
    output_excel = r"C:\data\output.xlsx"

    old_excel_file = r"C:\data\old_file.xlsx"
    new_excel_file = r"C:\data\new_file.xlsx"

    replace_links(input_excel, output_excel, old_excel_file, new_excel_file)
