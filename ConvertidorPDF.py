from tkinter import messagebox
import tkinter.font as tkFont
import win32com.client
import tkinter as tk
import pathlib
import os


def ppoint():
    # Objeto para la instancia
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    main_dir = pathlib.Path(os.getcwd())

    # Cambiar segun el formato pptx o ppt
    files = list(main_dir.glob("*.pptx"))
    for i, path in enumerate(files, start=1):
        output = path.with_suffix(".pdf")
        if output.exists():
            continue
        convertppoint(powerpoint, path, output)

    # Fin del proceso
    mensaje = " " + str(len(files)) + " archivos convertidos"
    messagebox.showinfo("Proceso completado", mensaje)


def convertppoint(powerpoint, path, output):
    document = powerpoint.Presentations.Open(str(path))
    document.SaveAs(output, 32)  # Parametro de 'formatType'
    document.Close()


def visio():
    # Objeto para la instancia
    visio = win32com.client.Dispatch("Visio.Application")
    visio.AlertResponse = 7
    main_dir = pathlib.Path(os.getcwd())

    # Enlista todos los documentos de MS Visio
    files = list(main_dir.glob("*.vsdx"))
    for i, path in enumerate(files, start=1):
        output = path.with_suffix(".pdf")
        if output.exists():
            continue
        convertvisio(visio, path, output)

    # Fin del proceso
    mensaje = " " + str(len(files)) + " archivos convertidos"
    messagebox.showinfo("Proceso completado", mensaje)


def convertvisio(visio, path, output):
    document = visio.Documents.Open(str(path))
    document.ExportAsFixedFormat(1, output, 1, 0)
    document.Close()


def main():
    # Instancias a mostrar
    ventana = tk.Tk()
    ventana.geometry("450x200")
    ventana.title("Creado por Oswaldo Tavares")
    estilo = tkFont.Font(family="Arial", size=20, weight="bold")
    txt1 = tk.Label(text="Convertidor de archivos", font=estilo, fg="orange")
    txt2 = tk.Label(text="   ")
    # Botones
    btn_visio = tk.Button(
        text="Visio a PDF", font=estilo, fg="blue", relief="groove", command=visio
    )
    txt3 = tk.Label(text="   ")
    btn_ppoint = tk.Button(
        text="Power Point a PDF", font=estilo, fg="red", relief="groove", command=ppoint
    )

    # Todo junto para ser mostrado
    txt1.pack()
    txt2.pack()
    btn_visio.pack()
    txt3.pack()
    btn_ppoint.pack()
    ventana.mainloop()


if __name__ == "__main__":
    main()
