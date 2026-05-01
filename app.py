import tkinter as tk

def generar_acta():
    nombre = entry_nombre.get()
    label_resultado.config(text=f"Acta generada para: {nombre}")

root = tk.Tk()
root.title("Sistema de Actas")

tk.Label(root, text="Nombre de la organización").pack()
entry_nombre = tk.Entry(root)
entry_nombre.pack()

tk.Button(root, text="Generar Acta", command=generar_acta).pack()

label_resultado = tk.Label(root, text="")
label_resultado.pack()

root.mainloop()