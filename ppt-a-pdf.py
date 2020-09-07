import win32com.client
import pathlib
import sys


def main():
    if len(sys.argv) < 2:
        print("¡¡Es necesario ingresar la ruta de la carpeta como argumento!!")
        sys.exit(1)

    # Objeto para la instancia
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    main_dir = pathlib.Path(sys.argv[1])

    # Cambiar segun el formato
    # pptx o ppt
    files = list(main_dir.glob("*.pptx"))
    print("\n| Total | Archivo")
    for i, path in enumerate(files, start=1):
        print(f"| {i} : {len(files)} | {path.stem}")
        output = path.with_suffix(".pdf")
        if output.exists():
            continue
        convert(powerpoint, path, output)


def convert(powerpoint, path, output):
    document = powerpoint.Presentations.Open(str(path))
    document.SaveAs(output, 32)  # Parametro de 'formatType'
    document.Close()


if __name__ == "__main__":
    main()
