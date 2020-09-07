import win32com.client
import pathlib
import sys


def main():
    if len(sys.argv) < 2:
        print("¡¡Es necesario ingresar la ruta de la carpeta como argumento!!")
        sys.exit(1)

    # Objeto para la instancia
    visio = win32com.client.Dispatch("Visio.Application")
    visio.AlertResponse = 7
    main_dir = pathlib.Path(sys.argv[1])

    # Enlista todos los documentos de MS Visio
    files = list(main_dir.glob("*.vsdx"))
    print("\n| Total | Archivo")
    for i, path in enumerate(files, start=1):
        print(f"| {i} : {len(files)} | {path.stem}")
        output = path.with_suffix(".pdf")
        if output.exists():
            continue
        convert(visio, path, output)


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
