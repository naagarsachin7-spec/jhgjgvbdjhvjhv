import argparse
from pathlib import Path

from openpyxl import load_workbook

from processor import fill_template, parse_invoice_file


def main():
    parser = argparse.ArgumentParser(description="Fill GST master template from invoice Excel files")
    parser.add_argument("invoices", nargs="+", help="Invoice Excel files")
    parser.add_argument("-t", "--template", default="template_clean.xlsx", help="Template workbook path")
    parser.add_argument("-o", "--output", default="GST_Master_Output.xlsx", help="Output workbook path")
    args = parser.parse_args()

    records = [parse_invoice_file(path) for path in args.invoices]
    wb = load_workbook(args.template)
    wb = fill_template(wb, records)
    wb.save(args.output)
    print(f"Saved: {args.output}")


if __name__ == "__main__":
    main()
