import tkinter as tk
import moduloScroll, openpyxl
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import *

def main(root, scrollable_frame):
    rutaGrupo = filedialog.askopenfilename(title="Elige un grupo de alumnos para editar",
                                                   initialdir="./_internal/Grupos/",
                                                   filetypes=[("Archivos xlsx", "*.xlsx")])

    frameGrupo = ttk.Frame(scrollable_frame)
    frameGrupo.pack()
    # Estilo de los widgets
    style = ttk.Style()
    style.map("TButton", foreground=[('pressed', 'blue'), ('active', 'black')],
              background=[('pressed', '!disabled', '#BFD4F3'), ('active', '#C7E5FC')])
    ttk.Style().configure("TFrame", background="#649BEF")
    ttk.Style().configure("TLabel", background="#649BEF", font=(18))
    ttk.Style().configure("TButton", background="#649BEF", foreground="white",
                           borderwidth=0, font=(18), weigth="blod", bordercolor="blue", weight='bold')

    
    ttk.Label(frameGrupo, text="Pulsa sobre el alumno para elimimarlo").grid(row=0, column=0, pady=10)

    def eliminaAlumno(fila, alumnosFrame):
        alumnosFrame.destroy()
        grupo = openpyxl.load_workbook(rutaGrupo)
        # Selecciona la hoja que desees procesar (por ejemplo, la primera hoja)
        hoja = grupo.active
        hoja['A' + str(fila)].value = None
        columna = hoja['A']
        valores = [celda.value for celda in columna]
        ultimo = len(valores)

        hoja['A' + str(ultimo)].value = None
        valores = [elemento for elemento in valores if elemento is not None]
        
        fila = 1
        for valor in valores:
            if valor != '':
               hoja.cell(row=fila, column=1, value=valor)
               fila +=1
        grupo.save(rutaGrupo)

        muestraAlumnos()
    
    def añadirAlumno(añadirEntry, alumnosFrame):
        grupo = openpyxl.load_workbook(rutaGrupo)
        # Selecciona la hoja que desees procesar (por ejemplo, la primera hoja)
        hoja = grupo.active
        columna = hoja['A']
        valores = [celda.value for celda in columna]
        ultimo = len(valores)
        alumno = añadirEntry.get()
        valores.append(alumno)
        valores.sort()      
        fila = 1
        for valor in valores:
            if valor != '':
               hoja.cell(row=fila, column=1, value=valor)
               fila +=1
        grupo.save(rutaGrupo)
        alumnosFrame.destroy()

        muestraAlumnos()

    def aviso(button, alumnosFrame):
        aviso = askyesno(title="Eliminar alumno", message="Seguro que desea eliminar a este alumno")
        if aviso == True:
            eliminaAlumno(button, alumnosFrame)
    
    
        
    def muestraAlumnos():
        alumnosFrame = ttk.Frame(frameGrupo)
        alumnosFrame.grid()    
        grupo = openpyxl.load_workbook(rutaGrupo)
        # Selecciona la hoja que desees procesar (por ejemplo, la primera hoja)
        hoja = grupo.active

        # Recorre la primera columna e imprime el contenido de cada celda
        i=1
        for fila in hoja.iter_rows(min_row=1, min_col=1, max_col=1, values_only=True):
            for celda in fila:
                if celda:
                    etiqueta = ttk.Label(alumnosFrame, text=f"{i}")
                    etiqueta.grid(row=i, column=0, padx=20)
                    button = i
                    botonEliminar = ttk.Button(alumnosFrame, text=f"Eliminar {celda}", 
                                               command=lambda button=button : aviso(button, alumnosFrame))
                    botonEliminar.grid(row=i, column=1, sticky="w")
                    i+=1

        añadirLabel = ttk.Label(alumnosFrame, text="Añadir alumno")
        añadirLabel.grid(row=i+1, column=0, padx=10, pady=20)
        añadirEntry = tk.Entry(alumnosFrame, width=40)
        añadirEntry.grid(row=i+1, column=1)
        añadirButton = ttk.Button(alumnosFrame, text="Añadir", command=lambda : añadirAlumno(añadirEntry, alumnosFrame))
        añadirButton.grid(row=i+1, column=3, padx=20)
        
    muestraAlumnos()
    

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Grupo de alumnos")
    root.geometry("700x1000")
    
    #---------- FRAME PARA METER EL CONTENIDO -------------------------------
    scrollable_frame = moduloScroll.ScrollableFrame(root)
    scrollable_frame.pack(fill="both", expand=True)

    main(root, scrollable_frame.scrollable_frame)

    root.mainloop()