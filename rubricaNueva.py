import tkinter as tk
import moduloScroll, openpyxl
from tkinter import ttk



def main(root, scrollable_frame):
    global entry_count, entry_count_row

    frameRubricaNueva = ttk.Frame(scrollable_frame)
    frameRubricaNueva.pack()
    
    # Crear los frames
    frameButton = ttk.Frame(frameRubricaNueva)
    frameButton.grid(row=0, padx=10, pady=10)
    frameLabel = ttk.Frame(frameRubricaNueva)
    frameLabel.grid(row=1, padx=10, pady=10)
    frame = ttk.Frame(frameRubricaNueva)
    frame.grid()
    row_frame = ttk.Frame(frame)
    row_frame.grid()

    inforNombre = ttk.Label(frameButton, text="Nombre de la rúbrica:")
    inforNombre.grid(row=2, column=0)
    nombreRubrica = ttk.Entry(frameButton)
    nombreRubrica.grid(row=2,column=1, padx=10, pady=5)
    nombreRubrica.configure(width=60)

    # Estilo de los widgets
    style = ttk.Style()
    style.map("TButton", foreground=[('pressed', 'blue'), ('active', 'black')],
              background=[('pressed', '!disabled', '#BFD4F3'), ('active', '#C7E5FC')])
    ttk.Style().configure("TFrame", background="#649BEF")
    ttk.Style().configure("TLabel", background="#649BEF", font=(18))
    ttk.Style().configure("TButton", background="#649BEF", foreground="white",
                           borderwidth=0, font=(18), weigth="blod", bordercolor="blue", weight='bold')
    
    # Variables
    entry_count = 4
    entry_count_row = 0

    def add_entry():
        if entry_count_row>0:
            add_button['state'] = 'disabled'
            add_button['text'] = 'Ahora no puedes añadir item'
            botonInfo = ttk.Label(frameLabel, text="Solo puedes añadir items cuando configuras la primera fila",
                                 fg="red")
            botonInfo.grid(row=2, padx=10, pady=5)
        else:
            global entry_count
            text = tk.Text(row_frame, height=2, width=30)
            text.grid(row=0, column=entry_count, padx=10, pady=5)
            entry_count += 1
    

    def add_row(): #ARREGLAR ESTO
        global entry_count_row
        entry_count_row +=1
        # Crear cuadros de texto dispuestos de forma horizontal en filas
        row_frame = ttk.Frame(frame)
        row_frame.grid(row=entry_count_row)

        for j in range(entry_count):  # Cambiar el número de cuadros de texto por fila según sea necesario
            text = tk.Text(row_frame, height=10, width=30)
            text.grid(row=0, column=j, padx=10, pady=5)


    def save_to_excel(frame):
        hoja = []
        wb = openpyxl.Workbook()
        hoja = wb.active
        rows = len(frame.winfo_children())
          
        # selecciona el texto de los Text
        for row in frame.winfo_children():    
            hoja.append([entry.get("1.0", "end-1c") for entry in row.winfo_children() if isinstance(entry,
             tk.Text)])
            
        print(hoja)
        # Guardar el array en un archivo Excel
        nombre = nombreRubrica.get()
        wb.save(f"./_internal/Rubricas/{nombre}.xlsx")



        save = tk.Label(frameLabel, text=f"La rúbrica se ha guardado correctamente en ./_internal/Rubricas/{nombre}.xlsx",
                        fg="green")
        save.grid(row=3, padx=10, pady=5)
    


#--------------Botones de la ventana ---------------------
    # Crear el botón para añadir más Entry
    add_button = ttk.Button(frameButton, text="Añadir item", command=add_entry)
    add_button.grid(row=0, column=0, padx=10, pady=5)

    # Crear fila de rúbricas
    add_fila = ttk.Button(frameButton, text="Añadir rúbrica", command=add_row)
    add_fila.grid(row=0, column=1, padx=10, pady=5)

    # Crear el botón para guardar en un archivo XLSX
    save_button = ttk.Button(frameButton, text="Guardar", command=lambda: save_to_excel(frame))
    save_button.grid(row=0, column=2, padx=10, pady=5)

    primeraFila = ("Criterios de evaluación", "Excelente", "Bien","Incompleto")
    for j in range(entry_count):  # Cambiar el número de cuadros de texto por fila según sea necesario
            text = tk.Text(row_frame, height=2, width=30)
            text.grid(row=0, column=j, padx=10, pady=5)
            print(primeraFila[j])
            text.insert(tk.END, primeraFila[j])

    
    criterio = tk.Label(frameLabel, text="En la primera casilla debes escribir 'Criterios de Evaluación'")
    criterio.grid(row=0, padx=10, pady=5)
    criterio.configure(background="#649BEF")
    info = tk.Label(frameLabel, text="En las siguientes casillas de la primera fila debes escribir "\
                    "cómo calificas la rúbrica. Excelente, bien, etc.")
    info.grid(row=1, padx=10, pady=5)
    info.configure(background="#649BEF")

    # Crear los cinco Entry iniciales
    for columna in range(entry_count_row):
        #global entry_count_row
        entry_count_row +=1
        # Crear cuadros de texto dispuestos de forma horizontal en filas
        row_frame = tk.Frame(frame)
        row_frame.grid()

        for j in range(entry_count):  # Cambiar el número de cuadros de texto por fila según sea necesario
            text = tk.Text(row_frame, height=10, width=30)
            text.grid(row=0, column=j, padx=10, pady=5)
            print(columna,j)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Grupo de alumnos")
    root.geometry("700x1000")
    
    #---------- FRAME PARA METER EL CONTENIDO -------------------------------
    scrollable_frame = moduloScroll.ScrollableFrame(root)
    scrollable_frame.pack(fill="both", expand=True)

    main(root, scrollable_frame.scrollable_frame)

    root.mainloop()
