import tkinter as tk
from typing import Protocol

TITLE = "BreastAGD282 Controller"


class Presenter(Protocol):
    def init_workbook_list(self) -> None: ...

    def handle_workbook_selected(self) -> None: ...

    def handle_calculation(self) -> None: ...


class MainWindow(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(TITLE)

    def init_ui(self, presenter: Presenter) -> None:
        self.frame = tk.Frame(self)
        self.workbook_selected = None  # set in init_workbook_list
        self.workbook_options = None  # set in init_workbook_list
        self.workbook_option_label = tk.Label(
            self.frame, text="Selected Excel workbook:"
        )
        presenter.init_workbook_list()
        self.workbook_selection = tk.StringVar(self.frame)
        self.workbook_selection.set(self.workbook_selected)
        self.workbook_selection.trace("w", presenter.handle_workbook_selected)
        self.workbook_option_menu = tk.OptionMenu(
            self.frame,
            self.workbook_selection,
            self.workbook_selected,
            *self.workbook_options,
            command=presenter.handle_workbook_selected
        )
        self.workbook_option_label.pack()
        self.workbook_option_menu.pack()

        self.calculate_button = tk.Button(
            self.frame, text="Calculate", command=presenter.handle_calculation
        )
        self.calculate_button.pack(side=tk.BOTTOM)

        self.frame.pack()

    def set_workbook_selection(self, value: str) -> None:
        self.workbook_selection.set(value)

    def init_workbook_list(self, active: str, options: list[str]) -> None:
        self.workbook_selected = active
        self.workbook_options = options

    def update_workbook_list(self, options: list[str]) -> None:
        self.workbook_option_menu["menu"].delete(0, "end")

        for book in options:
            self.workbook_option_menu["menu"].add_command(
                label=book,
                command=lambda value=book: self.workbook_selection.set(value),
            )
