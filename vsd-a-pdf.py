import win32com.client
import pathlib
import shutil
import sys
import os


def main():
    if len(sys.argv) < 2:
        print("¡¡Es necesario ingresar la ruta de la carpeta como argumento!!")
        sys.exit(1)

    # Open application
    visio = win32com.client.Dispatch("Visio.Application")

    visio.AlertResponse = 7

    main_dir = pathlib.Path(sys.argv[1])

    # All MS Visio documents
    files = list(main_dir.glob("*.vsdx"))
    print("\n| Total | Archivo")
    for i, path in enumerate(files, start=1):
        print(f"| {i} : {len(files)} | {path.stem}")
        output = path.with_suffix(".pdf")
        if output.exists():
            continue
        convert(visio, path, output)

    """
    # TO-DO
    main_dir = os.listdir("C:\\Users\\O.Tavares-ext\\Desktop\\VisioAPDF")
    output_f = "C:\\Users\\O.Tavares-ext\\Desktop\\New folder"

    print("")
    for files in main_dir:
        if files.endswith(".pdf"):
            print("Archivo con exito: " + files)
            shutil.move(files, output_f)

    print("\nProceso completado")
    """


def convert(visio, path, output):
    document = visio.Documents.Open(str(path))

    for page in document.Pages:
        # 0in esta dado por  pulgadas (formato interno)
        format_visio(page, "PageLeftMargin", "0in")
        format_visio(page, "PageRightMargin", "0in")
        format_visio(page, "PageTopMargin", "0in")
        format_visio(page, "PageBottomMargin", "0in")

    document.ExportAsFixedFormat(1, output, 1, 0)
    document.Close()


def format_visio(page, cell, value):
    page.PageSheet.Cells(cell).Formula = value


if __name__ == "__main__":
    main()
