import xlwings as xw
import logging
from datetime import datetime


def setup_logger():
    log_name = f"excel_link_replace_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logging.basicConfig(
        filename=log_name,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter("%(levelname)s - %(message)s")
    console.setFormatter(formatter)

    logging.getLogger().addHandler(console)

    return log_name


def replace_links(input_file, output_file, old_link, new_link):
    app = xw.App(visible=False)
    app.display_alerts = False

    changed_cells = 0
    changed_formulas = 0

    try:
        wb = xw.Book(input_file)
        logging.info(f"Opened file: {input_file}")

        for sheet in wb.sheets:
            logging.info(f"Processing sheet: {sheet.name}")

            rng = sheet.used_range

            for cell in rng:

                # ---------------- FORMELN ----------------
                try:
                    if cell.formula and isinstance(cell.formula, str):
                        if old_link in cell.formula:
                            cell.formula = cell.formula.replace(old_link, new_link)
                            changed_formulas += 1
                            logging.debug(f"Formula changed in {sheet.name}!{cell.address}")
                except Exception as e:
                    logging.warning(f"Formula error in {sheet.name}!{cell.address}: {e}")

                # ---------------- TEXTE ----------------
                try:
                    if cell.value and isinstance(cell.value, str):
                        if not (isinstance(cell.formula, str) and cell.formula.startswith("=")):
                            if old_link in cell.value:
                                cell.value = cell.value.replace(old_link, new_link)
                                changed_cells += 1
                                logging.debug(f"Value changed in {sheet.name}!{cell.address}")
                except Exception as e:
                    logging.warning(f"Value error in {sheet.name}!{cell.address}: {e}")

        # ---------------- WORKBOOK LINKS ----------------
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
                        logging.info(f"Workbook link replaced: {link}")
        except AttributeError:
            logging.info("No external workbook links found.")
        except Exception as e:
            logging.error(f"Error while processing workbook links: {e}")

        wb.save(output_file)
        wb.close()

        logging.info(f"Saved file: {output_file}")
        logging.info(f"Total formulas changed: {changed_formulas}")
        logging.info(f"Total cell values changed: {changed_cells}")

    except FileNotFoundError as e:
        logging.error(f"File not found: {e}")
        raise

    except Exception as e:
        logging.critical(f"Unexpected error: {e}")
        raise

    finally:
        app.quit()
        logging.info("Excel app closed.")


if __name__ == "__main__":
    log_file = setup_logger()
    logging.info(f"Log file: {log_file}")

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
