from tkinter import messagebox 
import tkinter.font as tkFont
import win32com.client
import tkinter as tk
import pathlib
import os

def convertppoint(powerpoint, path, output):
    document = powerpoint.Presentations.Open(str(path))
    document.SaveAs(output, 32)  # Parametro de 'formatType'
    document.Close()

def ppoint():
    # Objeto para la instancia
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    main_dir = pathlib.Path(os.getcwd())

    # Cambiar segun el formato
    # pptx o ppt
    files = list(main_dir.glob("*.pptx"))
    print("\n| Total | Archivo")
    for i, path in enumerate(files, start=1):

        # Progreso de archivo
        print(f"| {i} : {len(files)} | {path.stem}")
        output = path.with_suffix(".pdf")
        if output.exists():
            continue
        convertppoint(powerpoint, path, output)
    # TO-DO
    messagebox.showinfo("showinfo", "Proceso exitoso") 

def visio():
	# Objeto para la instancia
    visio = win32com.client.Dispatch("Visio.Application")
    visio.AlertResponse = 7
    main_dir = pathlib.Path(os.getcwd())

    # Enlista todos los documentos de MS Visio
    files = list(main_dir.glob("*.vsdx"))
    print("\n| Total | Archivo")
    for i, path in enumerate(files, start=1):
        print(f"| {i} : {len(files)} | {path.stem}")
        output = path.with_suffix(".pdf")
        if output.exists():
            continue
        convertvisio(visio, path, output)

def convertvisio(visio, path, output):
	document = visio.Documents.Open(str(path))
	document.ExportAsFixedFormat(1, output, 1, 0)
	document.Close()

def main():
    # Instancias a mostrar
    ventana = tk.Tk()
    ventana.geometry("450x250")
    ventana.title("Creado por Oswaldo Tavares")
    estilo = tkFont.Font(family="Arial", size=20)

    txt1 = tk.Label(text="Convertidor de archivos\n", font=estilo, fg="orange")

    # Botones
    btn_visio = tk.Button(text="Visio a PDF", font=estilo, command=visio)
    btn_ppoint = tk.Button(text="Power Point a PDF", font=estilo, command=ppoint)

    # Todo junto para ser mostrado
    txt1.pack()
    btn_visio.pack() # side=tk.LEFT
    btn_ppoint.pack() # width=50
    ventana.mainloop()

if __name__ == "__main__":
    main()
