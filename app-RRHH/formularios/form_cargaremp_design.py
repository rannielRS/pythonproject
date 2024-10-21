import tkinter as tk
from tkinter import *
import pymssql
import psycopg2
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR


class FormularioCargarEDesign():

    def __init__(self, panel_principal):   
       
        #Definiendo variables
        
        # Definiendo controles 

        #Tree
        #Boton para cargar empleados
        
        self.btn_cargar = tk.Button(panel_principal, text="Cargar empleados", font=(
            'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.cargartreeE)
        self.btn_cargar.grid(row=0,column=0,padx=0,pady=5)

        #Treeview
        
        
        columns = ('numero', 'nombreap', 'ci', 'escalas', 'thoraria', 'area', 'destajo')
        self.treeE = ttk.Treeview(panel_principal, height=25, columns=columns,
                                 show='headings')
        self.treeE.column('numero',width=80)
        self.treeE.column('nombreap',width=200)
        self.treeE.column('ci',width=130)
        self.treeE.column('escalas',width=100)
        self.treeE.column('thoraria',width=100)
        self.treeE.column('area',width=170)
        self.treeE.column('destajo',width=100)

        self.treeE.heading(column='numero', text='No.')
        self.treeE.heading(column='nombreap', text='Nombre y apellidos')
        self.treeE.heading(column='ci', text='CI')
        self.treeE.heading(column='escalas', text='Escala salarial')
        self.treeE.heading(column='thoraria', text='Tarifa horaria')
        self.treeE.heading(column='area', text='Área')
        self.treeE.heading(column='destajo', text='Destajo')
        self.treeE.grid(row=1,column=0, columnspan=5,ipadx=5,padx=5,pady=5)
        

        


    #Definiendo tree view de periodo
    def cargartreeE(self):
        pass
        # try:
        #     conn = pymssql.connect(
        #     server='10.105.213.6',
        #     user='userutil',
        #     password='1234',
        #     database='ZUNpr',
        #     as_dict=True
        #     )
        #     cursor = conn.cursor()
        #     if self.anno_trim:
        #         cursor.execute("SELECT * FROM ZUNpr.dbo.h_empleado WHERE ")
        #     else:
        #         cursor.execute('SELECT id_peri,nombre,fecha_inicio  FROM ZUNpr.dbo.n_periodo')
        #     slist = cursor.fetchall()
        #     for row in slist:
        #         options.append(str(row['id_peri'])+"-"+str(row['nombre']).rstrip()+"-"+str(row['fecha_inicio'])[:4])  

        #     self.cb_periodo['value']=options
            

        # except Exception:
        #     messagebox.showerror("Error","Problema de conexión con la base de datos")

        

        

        

    
    