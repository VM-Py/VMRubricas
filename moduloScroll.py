# modulo que devuelve un frame llamado scrollable_frame.scrollable_frame
# Lo que metamos en este frame tendrá sus barras de scroll funionales
import tkinter as tk
from tkinter import ttk

class ScrollableFrame(tk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)

        self.canvas = tk.Canvas(self)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

        self.scrollbar_vertical = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollbar_horizontal = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.scrollbar_vertical.set, xscrollcommand=self.scrollbar_horizontal.set)

        self.scrollbar_vertical.pack(side="right", fill="y")
        self.scrollbar_horizontal.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)

    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"), background="#649BEF")
    
    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.scrollable_frame, width=event.width)

# Ejemplo de uso
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Scrollable Frame Example")
    root.geometry("400x300")

    scrollable_frame = ScrollableFrame(root)
    scrollable_frame.pack(fill="both", expand=True)

    for i in range(30):
        tk.Label(scrollable_frame.scrollable_frame, text=f"Esto es un tecto de pruebas {i}").pack()

    root.mainloop()
