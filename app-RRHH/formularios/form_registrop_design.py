import tkinter as tk
from tkinter import ttk


class FormularioRegistroPDesign():

    def __init__(self, panel_principal):           
        # Crear dos subgráficos usando Matplotlib
        self.labelTitulo = tk.Label(panel_principal, text="Período")
        self.labelTitulo.grid(row=0,column=0,padx=20,pady=20)
        self.cb_periodo = ttk.Combobox(panel_principal)
        self.cb_periodo.grid(row=0,column=1,padx=5,pady=20)




    