from gerador_relatorios.gui import MedicalReportGenerator
import tkinter as tk


def main() -> None:
    root = tk.Tk()
    app = MedicalReportGenerator(root)
    root.mainloop()


if __name__ == "__main__":
    main()
