import re
import time
import os
import pathlib
import openpyxl
import sqlite3
import sys


def new_file_search() -> list[str]:
    return [_ for _ in os.listdir(pathlib.Path(__file__).parent.resolve()) if _.endswith("xlsx")]


def is_processing_row(row: tuple) -> bool:
    if all(_ is None for _ in row):
        return False
    if row[0] is not None and re.search(r"\d{2}/\d{2}/\d{2}", row[0]):
        return True
    return False


def file_etl(file_path: str, table_path: str):
    print(f"Processing: {file_path}")
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    for _row in filter(lambda x: is_processing_row(x), sheet.iter_rows(values_only=True)):
        #  sql = "INSERT INTO [bill] ([logDate], [description1], [description2], [hours]) " & _
        # 'values ("' & $reformatDate & '", "' & $activity & '", "' & $description & '", "' & $labor & '");'
        pass
    pass


def write_report(file_path: str, table_path: str):
    pass


def file_move_to_archive(file_path: str):
    print(f"Moving {file_path} to completed directory")
    pass


def create_sqlite_db(table_path: str):
    print("Initializing database")
    con = sqlite3.connect(table_path)
    _sql = (
        "CREATE TABLE [bill] ("
        "[logDate] DATE, "
        "[description1] TEXT, "
        "[description2] TEXT, "
        "[hours] decimal, "
        "[id] integer "
        "PRIMARY KEY AUTOINCREMENT);"
    )
    con.execute(_sql)
    con.close()


def process_all_files(processing_files: list[str]):
    sqlite_table = "report.db"
    if os.path.exists(sqlite_table):
        os.remove(sqlite_table)
    create_sqlite_db(sqlite_table)

    for _file in processing_files:
        file_etl(_file, sqlite_table)
        # write_report(file_path, table_path)
        # file_move_to_archive(file_path)

    # delete sqlite table
    # os.remove(sqlite_table)
    pass


def main():
    # search for files in import path ending in .xlxs
    # if not search, disply no report files message
    # if search, run processing
    # for each file:
    # create sqlite table, if not exists
    # clear all rows for sqlite table, and vaccuum
    # import file to sqlite table
    # run ETL?
    # export new file
    # move original file to comppleted directory
    # when all files done, delete sqlite table, if still exits
    # show processing complete message
    process_files = new_file_search()
    if not process_files:
        print("No files to process, exit")
        time.sleep(3)
        sys.exit()
    process_all_files(process_files)
    pass


if __name__ == "__main__":
    main()
