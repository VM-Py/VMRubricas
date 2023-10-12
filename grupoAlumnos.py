import tkinter as tk
import moduloScroll, openpyxl
from tkinter import ttk

def main(root, scrollable_frame):
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
    
    #--------- CONTENIDO DEL FRAME --------------------------------------------
    text_entries = []
    label = tk.Label(frameGrupo, text="Nombre del grupo")
    label.grid(row=0, column=0, padx=10, pady=5)
    label.configure(background="#649BEF")
    nombreGrupo = tk.Entry(frameGrupo, width=40)
    nombreGrupo.grid(row=0, column=1, padx=10, pady=5)

    def save_and_print():
        alumnos = []
        wb = openpyxl.Workbook()
        alumnos = wb.active
        for entry in text_entries:
            if entry !='':
                alumnos.append([entry.get()])
        
        # Ordena los alumnos alfabéticamente y los guarda
        columna = alumnos['A']
        valores = [celda.value for celda in columna]
        valores.sort()
        fila = 1
        for valor in valores:
            if valor != '':
               alumnos.cell(row=fila, column=1, value=valor)
               fila +=1

        nombre = nombreGrupo.get()
        wb.save(f"./_internal/Grupos/{nombre}.xlsx")

        print(f"Información guardada en el archivo /_internal/Grupos/{nombre}.xlsx")


    for i in range(25):
        label = tk.Label(frameGrupo, text=f"Alumno {i+1}:")
        label.grid(row=i+1, column=0, padx=10, pady=5)
        label.configure(background="#649BEF")
        text = tk.Entry(frameGrupo, width=50)
        text.grid(row=i+1, column=1, padx=10, pady=5)
        text_entries.append(text)
    button = ttk.Button(frameGrupo, text="Guardar", command=save_and_print)
    button.grid(row=0, column=2, padx=10, pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Grupo de alumnos")
    root.geometry("700x1000")
    
    #---------- FRAME PARA METER EL CONTENIDO -------------------------------
    scrollable_frame = moduloScroll.ScrollableFrame(root)
    scrollable_frame.pack(fill="both", expand=True)

    main(root, scrollable_frame.scrollable_frame)

    root.mainloop()

