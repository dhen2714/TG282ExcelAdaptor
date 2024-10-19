from view import MainWindow
from model import Model
from controller import Presenter


def main() -> None:
    model = Model("mammo_template.toml")
    view = MainWindow()
    presenter = Presenter(model, view)
    presenter.run()


if __name__ == "__main__":
    main()
