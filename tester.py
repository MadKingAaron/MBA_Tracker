import mba_database_parser
import load_and_save_file_dialog
if __name__ == "__main__":
    spread = mba_database_parser.Spreadsheet()
    spread.copyClipboardToAllWS()
    spread.saveWorkbook()