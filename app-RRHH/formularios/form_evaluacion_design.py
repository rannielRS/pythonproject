import tkinter as tk
from tkinter import *
import pymssql
import psycopg2
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR
from tkinter import tix


class FormularioEvaluacionDesign():

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
            self.btn_agEva = tk.Button(panel_principal, text="Registrar evaluación", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.registrarEva)
            self.btn_agEva.place(x=800, y=290)

            #Boton para agregar eva        
            self.btn_saveEva = tk.Button(panel_principal, text="Salvar evaluaciones", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.salvarEvaluacion)
            self.btn_saveEva.place(x=800, y=350)

            #Boton aprobar evaluaciones        
            self.btn_signEva = tk.Button(panel_principal, text="Revisar evaluaciones", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.signEva)
            self.btn_signEva.place(x=800, y=410)

            #label empleado
            self.lb_sempleado = tk.Label(panel_principal, text='Empleado seleccionado', justify='right', bg=COLOR_CUERPO_PRINCIPAL, font=('Times', 12), width=25)
            self.lb_sempleado.place(x=800, y=100)       
            
            #label mes1
            self.lb_mes1 = tk.Label(panel_principal, text='Mes 1', justify='right', bg=COLOR_CUERPO_PRINCIPAL, font=('Times', 12), width=10)
            self.lb_mes1.place(x=784, y=150)
            #label mes2
            self.lb_mes2 = tk.Label(panel_principal, text='Mes 2', justify='right', bg=COLOR_CUERPO_PRINCIPAL, font=('Times', 12), width=10)
            self.lb_mes2.place(x=784, y=200)
            #label mes3
            self.lb_mes3 = tk.Label(panel_principal, text='Mes 3', justify='right', bg=COLOR_CUERPO_PRINCIPAL, font=('Times', 12), width=10)
            self.lb_mes3.place(x=784, y=250) 

            #ComboEva x mes
            self.cb_mes1 = ttk.Combobox(panel_principal, postcommand=self.cargarCE, width=10)
            #self.cb_periodo.current(0)
            self.cb_mes1.place(x=870, y=150)

            self.cb_mes2 = ttk.Combobox(panel_principal, postcommand=self.cargarCE, width=10)
            #self.cb_periodo.current(0)
            self.cb_mes2.place(x=870, y=200)

            self.cb_mes3 = ttk.Combobox(panel_principal, postcommand=self.cargarCE, width=10)
            #self.cb_periodo.current(0)
            self.cb_mes3.place(x=870, y=250)       
            

            #Treeview        
            columns = ('numero', 'nombreap', 'ci', 'area','mes1','mes2','mes3')
            self.treeE = ttk.Treeview(panel_principal, height=16, columns=columns, show='headings')
            self.style = ttk.Style(self.treeE)
            self.style.configure('Treeview',rowheight=30)
            self.treeE.column('numero',width=80)
            self.treeE.column('nombreap',width=200)
            self.treeE.column('ci',width=110)
            self.treeE.column('area',width=200)
            self.treeE.column('mes1',width=60)
            self.treeE.column('mes2',width=60)
            self.treeE.column('mes3',width=60)

            self.treeE.heading(column='numero', text='No.')
            self.treeE.heading(column='nombreap', text='Nombre y apellidos')
            self.treeE.heading(column='ci', text='CI')
            self.treeE.heading(column='area', text='Área')
            slistP=self.getPeriodo()
            self.treeE.heading(column='mes1', text=slistP[0][1][0:3])
            self.treeE.heading(column='mes2', text=slistP[1][1][0:3])
            self.treeE.heading(column='mes3', text=slistP[2][1][0:3])
            self.treeE.grid(row=1,column=0, columnspan=5,ipadx=5,padx=5,pady=5)
            self.treeE.bind('<<TreeviewSelect>>',self.selectEmp)
            self.actualizarTreeE()  
            self.cargarCE()  
            self.cargarDpto()   
        else:
            messagebox.showinfo('Notificación','Debe registrar un período de evaluación')


    #Definiendo tree view de periodo
    def registrarEva(self):        
        curItem = self.treeE.focus()
        if curItem:   
            self.treeE.set(curItem, column='mes1', value=self.cb_mes1.get())         
            self.treeE.set(curItem, column='mes2', value=self.cb_mes2.get()) 
            self.treeE.set(curItem, column='mes3', value=self.cb_mes3.get()) 
        else:            
            messagebox.showinfo('Información','Debe seleccionar un trabajador')

        
    def actualizarTreeE(self):
        self.treeE.delete(*self.treeE.get_children())         
        queryEmpL=''
        if self.tx_empleado.get() != '' and self.cb_area.get() == '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id where x.nombreap like '%"+self.tx_empleado.get().upper()+"%'"
        elif self.cb_area.get() != '' and self.tx_empleado.get() == '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id where a.area = '"+self.cb_area.get()+"'"
        elif self.tx_empleado.get() != '' and self.cb_area.get() != '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id where x.nombreap like '%"+self.tx_empleado.get().upper()+"%' and a.area = '"+self.cb_area.get()+"'"
        else:
            queryEmpL='SELECT x.id,x.nombreap,x.ci,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id'
        
         
        self.cursorLoc.execute(queryEmpL)

        slistEmp = self.cursorLoc.fetchall()            
        for row in slistEmp:
            evames1=evames2=evames3='NE'          
            if self.obtenerEvaCond(row[0],self.getPeriodo()[0][0]) != None:
                evames1 = self.obtenerEvaCond(row[0],self.getPeriodo()[0][0])[0]                 
            if self.obtenerEvaCond(row[0],self.getPeriodo()[1][0]) != None:
                evames2 = self.obtenerEvaCond(row[0],self.getPeriodo()[1][0])[0]
            if self.obtenerEvaCond(row[0],self.getPeriodo()[2][0]) != None:
                evames3 = self.obtenerEvaCond(row[0],self.getPeriodo()[2][0])[0]

            self.treeE.insert('','end',values=("'"+row[0]+"'",row[1],row[2],row[3],evames1,evames2,evames3)) 


    def actualizarTreeE1(self,event):
        self.actualizarTreeE()

    def selectEmp(self,event):
        curItem = self.treeE.focus()
        selectedItem=self.treeE.item(curItem)
        self.cb_mes1.set(str(selectedItem['values'][4]))
        self.cb_mes2.set(str(selectedItem['values'][5]))
        self.cb_mes3.set(str(selectedItem['values'][6]))
        cadena=str(selectedItem['values'][0])+" "+str(selectedItem['values'][1])
        if len(cadena) <26:
            self.lb_sempleado['text']=cadena
        else:
            self.lb_sempleado['text']=cadena[0:20]+"..."
         
        queryP='SELECT * FROM postgres.public.periodo order by id asc'
        self.cursorLoc.execute(queryP)
        slistP=self.getPeriodo()
        self.lb_mes1['text']=slistP[0][1]+": "
        self.lb_mes2['text']=slistP[1][1]+": "
        self.lb_mes3['text']=slistP[2][1]+": "
        


    def cargarCE(self):
        options=[]        
        queryP='SELECT x.* FROM postgres.public.tipo_evaluacion x order by id asc'
        self.cursorLoc.execute(queryP)
        slistE=self.cursorLoc.fetchall()
        for row in slistE:
            options.append(row[1])
        
        self.cb_mes1['values']=options
        self.cb_mes2['values']=options
        self.cb_mes3['values']=options

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

    def obtenerTipoEva(self,eva):
        queryP="SELECT id FROM postgres.public.tipo_evaluacion where eva='"+str(eva)+"'"
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()

    def obtenerEvaCond(self,emp,periodo):
        queryP="SELECT te.eva FROM postgres.public.evaluacion AS e  INNER JOIN postgres.public.tipo_evaluacion AS te ON e.evaluacion_tipoevaluacion_id = te.id where e.evaluacion_empleado_id='"+str(emp)+"' and e.evaluacion_perio_id='"+str(periodo)+"'"
        #print(queryP)
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchone()
    
    def salvarEvaluacion(self):

        for parent in self.treeE.get_children():
            connLoc = psycopg2.connect(
            host="localhost",
            database="postgres",
            user="postgres",
            password="proyecto") 
            cursorLoc = connLoc.cursor() 
            #cursorLoc.execute(queryP)
            #return cursorLoc.fetchall()
            listTree=self.treeE.item(parent)["values"]
            slistP=self.getPeriodo()
            queryInEvaM1="INSERT INTO postgres.public.evaluacion (evaluacion_empleado_id,evaluacion_tipoevaluacion_id,evaluacion_perio_id) \
                        VALUES ("+listTree[0]+","+str(self.obtenerTipoEva(listTree[4])[0][0])+","+str(slistP[0][0])+")"
            queryInEvaM2="INSERT INTO postgres.public.evaluacion (evaluacion_empleado_id,evaluacion_tipoevaluacion_id,evaluacion_perio_id) \
                        VALUES ("+listTree[0]+","+str(self.obtenerTipoEva(listTree[5])[0][0])+","+str(slistP[1][0])+")"
            queryInEvaM3="INSERT INTO postgres.public.evaluacion (evaluacion_empleado_id,evaluacion_tipoevaluacion_id,evaluacion_perio_id) \
                        VALUES ("+listTree[0]+","+str(self.obtenerTipoEva(listTree[6])[0][0])+","+str(slistP[2][0])+")"
            self.cursorLoc.execute(queryInEvaM1)
            self.cursorLoc.execute(queryInEvaM2)
            self.cursorLoc.execute(queryInEvaM3)
            self.connLoc.commit()

    def signEva(self):
        pass