import tkinter as tk
from tkinter import *
import pymssql
import psycopg2
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR
import openpyxl

class FormularioCalcUtilidadesDesign():

    def __init__(self, panel_principal):   
       
        #Definiendo variables
        # Variables de conexion
        self.connLoc = psycopg2.connect(
            host="localhost",
            database="postgres",
            user="postgres",
            password="proyecto") 
        self.cursorLoc = self.connLoc.cursor()
        if self.getPeriodo():
            
            # Definiendo controles de seleccion
            self.tx_empleado = ttk.Entry(panel_principal, font=('Times', 14), width=10)
            self.tx_empleado.grid(row=0,column=0,padx=5,pady=5,ipadx=40)

            #Boton para buscar empleados        
            self.btn_bempleados = tk.Button(panel_principal, text="Buscar", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.actualizarTreeE)
            self.btn_bempleados.place(x=250, y=2)
            #Para buscar por departamento
            #Label area
            self.tx_area = tk.Label(panel_principal, font=('Times', 14), width=20, bg=COLOR_CUERPO_PRINCIPAL, text='Departamento:')
            self.tx_area.place(x=350, y=5)

            #Combo departamento
            self.cb_area= ttk.Combobox(panel_principal, width=30)
            self.cb_area.bind('<<ComboboxSelected>>', self.actualizarTreeE1)
            #self.cb_periodo.current(0)
            self.cb_area.place(x=520, y=5)

            #Boton para agregar eva        
            self.btn_agEva = tk.Button(panel_principal, text="Registrar salario", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.regSal)
            self.btn_agEva.place(x=800, y=290)

            #Boton para agregar eva        
            self.btn_saveEva = tk.Button(panel_principal, text="Registrar vacaciones", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.regVac)
            self.btn_saveEva.place(x=800, y=350)

            #Boton aprobar evaluaciones        
            self.btn_signEva = tk.Button(panel_principal, text="Mostrar resumen", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.showResumen)
            self.btn_signEva.place(x=800, y=410)

                  
            

            #Treeview        
            columns = ('numero', 'nombreap', 'ci', 'salario','vacaciones','horast','coef','devengado')
            self.treeE = ttk.Treeview(panel_principal, height=16, columns=columns, show='headings')
            self.style = ttk.Style(self.treeE)
            self.style.configure('Treeview',rowheight=30)
            self.treeE.column('numero',width=80)
            self.treeE.column('nombreap',width=200)
            self.treeE.column('ci',width=110)
            self.treeE.column('salario',width=60)
            self.treeE.column('vacaciones',width=60)
            self.treeE.column('horast',width=60)
            self.treeE.column('coef',width=60)
            self.treeE.column('devengado',width=80)

            self.treeE.heading(column='numero', text='No.')
            self.treeE.heading(column='nombreap', text='Nombre y apellidos')
            self.treeE.heading(column='ci', text='CI')
            self.treeE.heading(column='salario', text='Mt. Sal')
            self.treeE.heading(column='vacaciones', text='Mt. Vac')
            self.treeE.heading(column='horast', text='Horas T.')
            self.treeE.heading(column='coef', text='C. Eva')
            self.treeE.heading(column='devengado', text='S. Dev')
            self.treeE.grid(row=1,column=0, columnspan=5,ipadx=5,padx=5,pady=5)
            self.actualizarTreeE() 
            self.cargarDpto()   
        else:
            messagebox.showinfo('Notificación','Debe registrar un período de evaluación')


    #Definiendo tree view de periodo
    def regSal(self):        
        pass

        
    def actualizarTreeE(self):
        pass


    def actualizarTreeE1(self,event):
        self.actualizarTreeE()       



    def cargarDpto(self):
        options=[]         
        queryP='SELECT x.* FROM postgres.public.area x order by area asc'
        self.cursorLoc.execute(queryP)
        slistArea=self.cursorLoc.fetchall()
        for row in slistArea:
            options.append(row[1])
        
        self.cb_area['values']=options

    def getPeriodo(self):         
        queryP='SELECT p.* FROM postgres.public.utilidades_periodo_incluye x INNER JOIN postgres.public.periodo AS p ON x.upincluye_periodo_id = p.id order by p.id asc'
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()

    def obtenerPerMes(self,mes):
        queryP="SELECT * FROM postgres.public.periodo where mes='"+str(mes)+"'"
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()
   
    
    def regVac(self):
        pass

    

    def getDepartamento(self,idemp):         
        queryP="SELECT a.area  FROM postgres.public.empleado emp INNER JOIN postgres.public.area AS a ON emp.empleado_area_id  = a.id where emp.id = "+str(idemp)
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchone()[0]

    def showResumen(self):         
        pass

    
