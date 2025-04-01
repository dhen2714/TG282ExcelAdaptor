from excel_processing import (
    read_excel_parameter_toml,
    read_testing_spreadsheet,
    values_to_csvline,
)
import xlwings as xw
import csv
import subprocess

TMP_PATH = "tmp.csv"


class Model:
    def __init__(self, toml_path) -> None:
        self.active_sheet = None
        self.active_cell = None
        self.toml_dict = read_excel_parameter_toml(toml_path)
        self.path_to_executable = self.toml_dict.pop("path_to_executable")
        try:
            self.selected_book = xw.books.active
            self.selected_book_name = self.selected_book.name
        except xw.XlwingsError:
            self.selected_book = None
            self.selected_book_name = "-"

    def run_calculation(self) -> None:
        toml_dict = self.toml_dict.copy()
        if self.selected_book:
            print(f"Calculating MGD using values in {self.selected_book.name}.")
            spreadsheet_parameters = read_testing_spreadsheet(
                self.selected_book, toml_dict
            )
            if not spreadsheet_parameters:
                print("No data for calculation.")
                return

            csv_lines = {}
            for exposure_mode, values in spreadsheet_parameters.items():
                csv_line = values_to_csvline(values)
                csv_lines[exposure_mode] = csv_line

            with open(TMP_PATH, "w", newline="") as csvfile:
                csvwriter = csv.writer(csvfile)
                for exposure_mode, lineval in csv_lines.items():
                    line_list = lineval.value_list
                    csvwriter.writerow(line_list)

        process = subprocess.Popen(
            [self.path_to_executable, "tmp.csv"], stdin=subprocess.PIPE, text=True
        )
        process.communicate("\n")

    def get_book_names(self) -> list[str]:
        try:
            book_names = ["-"]
            for book in xw.books:
                if book.name != self.selected_book_name:
                    book_names.append(book.name)
            return book_names
        except Exception as e:
            self.selected_book = None
            self.selected_book_name = "-"
            return ["-"]

    def select_book(self, book_name: str) -> None:
        try:
            self.selected_book = xw.books[book_name]
            self.selected_book_name = self.selected_book.name
        except (KeyError, xw.XlwingsError):
            self.selected_book = None
            self.selected_book_name = "-"
