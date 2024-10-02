import tkinter as tk
from tkinter import ttk
from config import COLOR_CUERPO_PRINCIPAL


class FormularioRegistroPDesign():

    def __init__(self, panel_principal):           
        # Combo para cargar los periodos del Zun
        self.marco_trabajo = tk.Frame(
            panel_principal, bg=COLOR_CUERPO_PRINCIPAL, height=50)
        self.marco_trabajo.pack(side=tk.TOP,  fill='both')

        self.labelTitulo = tk.Label(self.marco_trabajo, text="Período", bg=COLOR_CUERPO_PRINCIPAL)
        self.labelTitulo.grid(row=0,column=0,padx=20,pady=20)

        self.cb_periodo = ttk.Combobox(self.marco_trabajo)
        self.cb_periodo.grid(row=0,column=1,padx=5,pady=20)        
        

        self.tx_trimestre_name = ttk.Entry(self.marco_trabajo, font=(
            'Times', 14), state="readonly", width=20)
        self.tx_trimestre_name.grid(row=0,column=2,padx=5,pady=20)

        self.treeview = ttk.Treeview(columns=("size", "lastmod"))
        self.treeview.heading("#0", text="Mes")
        self.treeview.heading("size", text="Periodo")
        self.treeview.heading("lastmod", text="Última modificación")



    