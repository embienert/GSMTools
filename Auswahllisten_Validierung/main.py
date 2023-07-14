from GSMTools.excel import Data, Reference

from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter import Tk


def check_data(data: Data, reference: Reference):
    issues = data.cross_reference(reference)
    # issues = data.cross_reference(reference, subset_title=["wgk", "reinheit"])  # Check nur auf Spalten Titel in der Liste
    # issues = data.cross_reference(reference, subset_header=["WGK", "Reinheit"])  # Check nur auf Spalten mit Header aus der Liste

    for column, discrepancies in issues.items():
        column_header = column.header.replace("\n", " ").strip()
        print(f"\nProbleme mit Spalte [{column_header}]:")

        for row_nr, text in discrepancies:
            if text in [None, "None"]:
                text = "Kein Eintrag vorhanden"
            print(f"\tZeile {row_nr}: \t{text}")


def create_data_validations(data: Data, reference: Reference):
    root = Tk()
    save_filename = asksaveasfilename(title="Save copy...", filetypes=[("Microsoft Excel File", "*.xlsx")])
    root.destroy()

    data.add_data_validations(reference)
    data.save_copy(save_filename)


def main():
    root = Tk()
    reference_filename = askopenfilename(title="Open reference file...", filetypes=[("Microsoft Excel File", "*.xlsx")])
    data_filename = askopenfilename(title="Open data file...", filetypes=[("Microsoft Excel File", "*.xlsx")])
    root.destroy()

    data = Data(data_filename)
    reference = Reference(reference_filename)

    check_data(data, reference)
    # create_data_validations(data, reference)


if __name__ == '__main__':
    main()
