import tkinter as tk
import pymssql
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR


class FormularioRegistroPDesign():

    def __init__(self, panel_principal):           
        # Definiendo controles 
        self.cbx_label = tk.Label(panel_principal, text="Período", bg=COLOR_CUERPO_PRINCIPAL)
        self.cbx_label.grid(row=1,column=0,padx=5,pady=20)
        #comobo carga periodo del zun
        self.cb_periodo = ttk.Combobox(panel_principal)
        self.cb_periodo.grid(row=1,column=1,padx=5,pady=20)        
        
        self.tx_label = tk.Label(panel_principal, text="Trimestre", bg=COLOR_CUERPO_PRINCIPAL)
        self.tx_label.grid(row=0,column=0,padx=5,pady=20)
        
        self.tx_trimestre_name = ttk.Entry(panel_principal, font=(
            'Times', 14), width=20)
        self.tx_trimestre_name.grid(row=0,column=1,padx=5,pady=20)

        #Boton para agregar periodo al treewiew
        self.btn_registro_periodo = tk.Button(panel_principal, text="Registrar", font=(
            'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, padx=15, command="")
        self.btn_registro_periodo.grid(row=1,column=2,padx=5,pady=20)
        #self.btn_registro_periodo.bind(
        #    "<Return>", (lambda event: self.registrar_producto()))


        #Definiendo tree view de periodo

        self.tree = ttk.Treeview(panel_principal,
                                 show='headings')
        self.tree['columns'] = ('Id', 'Periodo', 'Orden')
        self.tree.column('Id')
        self.tree.column('Periodo')
        self.tree.column('Orden')

        self.tree.heading('Id', text='Id')
        self.tree.heading('Periodo', text='Período')
        self.tree.heading('Orden', text='Orden')
        self.tree.grid(row=2,column=0, columnspan=4,padx=20,pady=20)

        self.combo()

    def combo(self):   
        try:
            conn = pymssql.connect(
            server='10.105.213.6',
            user='userutil',
            password='1234',
            database='ZUNpr',
            as_dict=True
            )
            cursor = conn.cursor()
            # cursor.execute('SELECT * FROM usuario')
            # slist = cursor.fetchall()
            # self.cb_periodo.configure(values=slist)
        except Exception as err:
            messagebox.showerror("Error",err)



    