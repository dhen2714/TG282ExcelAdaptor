import xlwings as xw
from dataclasses import dataclass, asdict
from pathlib import Path
import tomllib
from errors import SheetNotFound, ValueParseError


@dataclass
class CSVLine:
    view: str  # CC or MLO
    thickness: int  # in mm, 15 - 120
    density_percentage: int  # only when "none" selected as model
    density_percentile: int  # 5, 25, 50, 75 or 95
    modality: str  # DM, DBT, CEDM or CEDBT
    number_angles: int  # >= 1
    angle_range: float  # decimal 0.0 to 60.0
    source_to_dosimeter: int  # in mm
    source_to_table: int  # in mm
    anode: str  # Mo, Rh or W
    voltage: int  # 20 - 49
    filter_element: str  # chemical symbol
    filter_thickness: float  # in mm
    hvl: float  # > 0.0
    air_kerma: float  # measured air kerma
    high_spectrum_anode: str = ""  # These are for contrast enhanced views
    high_spectrum_voltage: int = ""
    high_spectrum_filter_element: str = ""
    high_spectrum_filter_thickness: float = ""
    high_spectrum_hvl: float = ""
    high_spectrum_air_kerma: float = ""
    dose_distribution: float = ""

    @property
    def value_list(self) -> list:
        return list(asdict(self).values())


def read_excel_parameter_toml(fpath: Path | str) -> dict:
    with open(fpath, "rb") as f:
        return tomllib.load(f)


def read_excel_addresses(xw_sheet: xw.main.Sheet, address_list: list[str]) -> dict:
    address_to_value = {}
    for address in address_list:
        address_to_value[address] = xw_sheet[address].value
    return address_to_value


def get_exposure_mode_values(excel_address_dict: dict, excel_sheet: xw.Sheet) -> dict:
    parameter_values = {}
    parameter_values["view"] = excel_address_dict.pop("view")
    parameter_values["modality"] = excel_address_dict.pop("modality")
    parameter_values["number_angles"] = excel_address_dict.pop("number_angles")
    parameter_values["angle_range"] = excel_address_dict.pop("angle_range")
    parameter_values["filter_thickness"] = excel_address_dict.pop("filter_thickness")
    parameter_values["density_percentage"] = excel_address_dict.pop(
        "density_percentage"
    )
    parameter_values["density_percentile"] = excel_address_dict.pop(
        "density_percentile"
    )

    for parameter, address in excel_address_dict.items():
        try:
            sheet_value = excel_sheet[address].value
        except Exception:
            raise ValueParseError(f"Error when trying to parse {parameter}: {address}")
        else:
            if sheet_value:
                parameter_values[parameter] = sheet_value
            else:
                raise ValueParseError(f"No {parameter} value found.")

    return parameter_values


def read_testing_spreadsheet(excel_book: xw.main.Book, toml_dict: dict) -> dict:
    sheet_name = toml_dict.pop("sheet")
    source_to_table_address = toml_dict.pop("source_to_breast_support")
    try:
        xlsheet = excel_book.sheets[sheet_name]
    except Exception:
        raise SheetNotFound(f"Could not find Excel sheet named {sheet_name}")

    spreadsheet_values = {}
    for exposure_mode, value_addresses in toml_dict.items():
        try:
            parameter_values = get_exposure_mode_values(value_addresses.copy(), xlsheet)
            parameter_values["source_to_table"] = int(
                xlsheet[source_to_table_address].value
            )
            source_to_dosimeter = int(parameter_values["source_to_table"]) - int(
                parameter_values["detector_above_support"]
            )
            parameter_values["source_to_dosimeter"] = int(source_to_dosimeter)
            spreadsheet_values[exposure_mode] = parameter_values

        except Exception as e:
            print(f"Could not parse values for {exposure_mode}.", e)

    return spreadsheet_values


def values_to_csvline(value_dict: dict[str]) -> CSVLine:
    return CSVLine(
        view=value_dict["view"],
        thickness=int(value_dict["thickness"]),
        density_percentage=(
            ""
            if value_dict["density_percentage"] == ""
            else int(value_dict["density_percentage"])
        ),
        density_percentile=(
            ""
            if value_dict["density_percentile"] == ""
            else int(value_dict["density_percentile"])
        ),
        modality=value_dict["modality"],
        number_angles=value_dict["number_angles"],
        angle_range=value_dict["angle_range"],
        source_to_dosimeter=value_dict["source_to_dosimeter"],
        source_to_table=value_dict["source_to_table"],
        anode=value_dict["anode"],
        voltage=int(value_dict["voltage"]),
        filter_element=value_dict["filter_element"],
        filter_thickness=value_dict["filter_thickness"],
        hvl=value_dict["hvl"],
        air_kerma=value_dict["air_kerma"],
    )


if __name__ == "__main__":
    test_path = "test_files/test.xls"
    testbook = xw.Book(test_path)
    toml_dict = read_excel_parameter_toml("mammo_template.toml")
    test_dict = read_testing_spreadsheet(testbook, toml_dict)
    csv_lines = {}
    for exposure_mode, values in test_dict.items():
        print(exposure_mode)
        csv_lines[exposure_mode] = values_to_csvline(values)

    import csv

    with open("../BreastAGD282_Windows/test.csv", "w", newline="") as csvfile:
        csvwriter = csv.writer(csvfile)
        for exposure_mode, lineval in csv_lines.items():
            line_list = list(asdict(lineval).values())
            csvwriter.writerow(line_list)

    test_csvline = values_to_csvline(test_dict["ACR_phantom"])

    import subprocess

    # subprocess.run(["BreastAGD282.exe", "test.csv"])
    # process = subprocess.Popen(
    #     ["BreastAGD282.exe", "test.csv"], stdin=subprocess.PIPE, text=True
    # )
    # process.communicate("\n")
