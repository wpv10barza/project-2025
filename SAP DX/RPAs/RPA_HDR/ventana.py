import tkinter as tk
def pantalla_final():
  ventana = tk.Tk()
  ventana.title("Finalización")
  ventana.geometry("500x300")

  ventana.attributes("-topmost",True)
  ventana.lift()

  mensaje = tk.Label(ventana, text="Proceso finalizado.", font=("Arial", 14))
  mensaje.pack(pady=20)

  boton = tk.Button(ventana, text="Cerrar", command=ventana.destroy)
  boton.pack()
  ventana.mainloop()
  
