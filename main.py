from view import MainWindow
from model import Model
from controller import Controller


def main() -> None:
    model = Model("mammo_template.toml")
    view = MainWindow()
    controller = Controller(model, view)
    controller.run()


if __name__ == "__main__":
    main()
