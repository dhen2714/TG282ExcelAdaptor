from view import MainWindow
from model import Model


class Controller:
    def __init__(self, model: Model, view: MainWindow) -> None:
        self.model = model
        self.view = view

    def run(self) -> None:
        self.view.init_ui(self)
        self.view.mainloop()

    def init_workbook_list(self) -> None:
        selected = self.model.selected_book_name
        try:
            book_names = self.model.get_book_names()
        except Exception as e:
            book_names = ["-"]
        self.view.init_workbook_list(selected, book_names)
        self.view.after(2000, self.update_workbook_list)

    def update_workbook_list(self) -> None:
        try:
            book_names = self.model.get_book_names()
            self.view.update_workbook_list(book_names)
        except Exception as e:
            self.view.update_workbook_list([])
            self.view.set_workbook_selection("-")
        self.view.after(2000, self.update_workbook_list)

    def handle_workbook_selected(self, *args) -> None:
        self.model.select_book(self.view.workbook_selection.get())

    def handle_calculation(self) -> None:
        self.model.run_calculation()
