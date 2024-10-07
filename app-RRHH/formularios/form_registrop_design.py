import tkinter as tk
from tkinter import *
import pymssql
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR


class FormularioRegistroPDesign():

    def __init__(self, panel_principal):           
        #Definiendo variables
        self.var_periodo = StringVar()
        self.anno_trim = StringVar()
        
        # Definiendo controles 

        #comobo carga periodo del zun
        self.cbx_label = tk.Label(panel_principal, text="Período", bg=COLOR_CUERPO_PRINCIPAL)
        self.cbx_label.grid(row=1,column=0,padx=5,pady=20)
        
        self.cb_periodo = ttk.Combobox(panel_principal,textvariable=self.var_periodo, postcommand=self.cargarcombo)
        self.cb_periodo.grid(row=1,column=1,padx=5,pady=20)  
        

        #comobo carga de orden
        self.cbx_labelO = tk.Label(panel_principal, text="Orden", bg=COLOR_CUERPO_PRINCIPAL)
        self.cbx_labelO.grid(row=1,column=2,padx=5,pady=20)
        
        self.cb_orden = ttk.Combobox(panel_principal,postcommand=self.cargarcombo,values=['1','2','3'])
        self.cb_orden.current(0)
        self.cb_orden.grid(row=1,column=3,padx=5,pady=20) 

        #nombre del periodo
        self.tx_label = tk.Label(panel_principal, text="Trimestre", bg=COLOR_CUERPO_PRINCIPAL)
        self.tx_label.grid(row=0,column=0,padx=5,pady=20)
        
        self.tx_trimestre_name = ttk.Entry(panel_principal, font=(
            'Times', 14), width=20)
        self.tx_trimestre_name.grid(row=0,column=1,padx=5,pady=20)

        #Año del periodo
        self.tx_labelA = tk.Label(panel_principal, text="Año", bg=COLOR_CUERPO_PRINCIPAL)
        self.tx_labelA.grid(row=0,column=2,padx=5,pady=20)
        
        self.annot = str(self.var_periodo)[:4]

        self.tx_anno = ttk.Entry(panel_principal, font=(
            'Times', 14), width=20)
        self.tx_anno.grid(row=0,column=3,padx=5,pady=20)

        #Boton para agregar periodo al treewiew
        self.btn_registro_periodo = tk.Button(panel_principal, text="Registrar", font=(
            'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, padx=15, command="")
        self.btn_registro_periodo.grid(row=1,column=4,padx=5,pady=20)
        


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

        

    def cargarcombo(self): 
        self.anno_trim=self.tx_anno.get()  
        
        options=[]
        try:
            conn = pymssql.connect(
            server='10.105.213.6',
            user='userutil',
            password='1234',
            database='ZUNpr',
            as_dict=True
            )
            cursor = conn.cursor()
            if self.anno_trim:
                cursor.execute("SELECT id_peri,nombre,fecha_inicio  FROM ZUNpr.dbo.n_periodo where fecha_inicio like '%"+self.anno_trim+"%'")
            else:
                cursor.execute('SELECT id_peri,nombre,fecha_inicio  FROM ZUNpr.dbo.n_periodo')
            slist = cursor.fetchall()
            for row in slist:
                options.append(str(row['id_peri'])+"-"+str(row['nombre']).rstrip()+"-"+str(row['fecha_inicio'])[:4])  

            self.cb_periodo['value']=options

        except Exception as err:
            messagebox.showerror("Error",err)

    

    

    