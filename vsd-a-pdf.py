import pathlib
import sys
import win32com.client


def set_ps(page, cell, value):
    page.PageSheet.Cells(cell).Formula = value


def convert(visio, path, out_path):
    doc = visio.Documents.Open(str(path))

    for page in doc.Pages:
        set_ps(page, "PageLeftMargin", "0mm")
        set_ps(page, "PageRightMargin", "0mm")
        set_ps(page, "PageTopMargin", "0mm")
        set_ps(page, "PageBottomMargin", "0mm")
        page.ResizeToFitContents()

    pdf_format = 1
    intent_print = 1
    print_all = 0
    doc.ExportAsFixedFormat(pdf_format, out_path, intent_print, print_all)
    doc.Close()


def main():
    visio = win32com.client.Dispatch("Visio.Application")
    visio.AlertResponse = 7  # Answer "no" to all save dialogs

    basedir = pathlib.Path(sys.argv[1])
    files = list(basedir.glob("*.vsdx"))
    for i, path in enumerate(files, start=1):
        print(f"[{i:3}/{len(files):3}] {path.stem}")
        out_path = path.with_suffix(path.suffix + ".pdf")
        if out_path.exists():
            continue
        convert(visio, path, out_path)


if __name__ == "__main__":
    main()
