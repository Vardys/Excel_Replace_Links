import xlwings as xw
import logging


def safe_replace(v, old, new):
    if isinstance(v, str) and old in v:
        return v.replace(old, new)
    return v


def replace_links_fast(input_file, output_file, old_link, new_link):
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    changed_formulas = 0
    changed_values = 0

    try:
        wb = xw.Book(input_file)
        logging.info("Workbook loaded")

        for sheet in wb.sheets:
            logging.info(f"Sheet: {sheet.name}")

            # -------------------------
            # 1. FORMELN in einem Block
            # -------------------------
            try:
                rng = sheet.api.UsedRange.SpecialCells(11)  # xlCellTypeFormulas
                formulas = rng.Formula

                if isinstance(formulas, tuple):
                    formulas = list(map(list, formulas))

                    for i in range(len(formulas)):
                        for j in range(len(formulas[i])):
                            if isinstance(formulas[i][j], str) and old_link in formulas[i][j]:
                                formulas[i][j] = formulas[i][j].replace(old_link, new_link)
                                changed_formulas += 1

                    rng.Formula = formulas

                elif isinstance(formulas, str):
                    if old_link in formulas:
                        rng.Formula = formulas.replace(old_link, new_link)
                        changed_formulas += 1

            except Exception:
                # keine Formeln im Sheet
                pass

            # -------------------------
            # 2. VALUES (nur konstante Werte)
            # -------------------------
            try:
                rng = sheet.api.UsedRange.SpecialCells(2)  # xlCellTypeConstants
                values = rng.Value

                if isinstance(values, tuple):
                    values = list(map(list, values))

                    for i in range(len(values)):
                        for j in range(len(values[i])):
                            if isinstance(values[i][j], str) and old_link in values[i][j]:
                                values[i][j] = values[i][j].replace(old_link, new_link)
                                changed_values += 1

                    rng.Value = values

                elif isinstance(values, str):
                    if old_link in values:
                        rng.Value = values.replace(old_link, new_link)
                        changed_values += 1

            except Exception:
                # keine Konstanten im Sheet
                pass

        # -------------------------
        # Workbook links (1 call)
        # -------------------------
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

        logging.info(f"Done. Formulas: {changed_formulas}, Values: {changed_values}")

    finally:
        app.calculation = "automatic"
        app.screen_updating = True
        app.quit()
