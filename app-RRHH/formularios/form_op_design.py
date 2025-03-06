import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR,CONN_LOC,CURSOR_LOC
import openpyxl
import os
import subprocess
class FormularioOtrosPagosDesign():

    def __init__(self, panel_principal):   
       
        #Definiendo variables
        # Variables de conexion
        self.connLoc = CONN_LOC
        self.cursorLoc = CURSOR_LOC
        if self.getPeriodo():
            
            # Definiendo controles de seleccion
            self.tx_empleado_op = ttk.Entry(panel_principal, font=('Times', 14), width=10)
            self.tx_empleado_op.grid(row=0,column=0,padx=5,pady=5,ipadx=40)

            #Boton para buscar empleados        
            self.btn_bempleados_op = tk.Button(panel_principal, text="Buscar", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.actualizartreeEOP)
            self.btn_bempleados_op.place(x=250, y=2)
            #Para buscar por departamento
            #Label area
            self.tx_area = tk.Label(panel_principal, font=('Times', 14), width=20, bg=COLOR_CUERPO_PRINCIPAL, text='Departamento:')
            self.tx_area.place(x=350, y=5)

            #Combo departamento
            self.cb_area_op= ttk.Combobox(panel_principal, width=30)
            self.cb_area_op.bind('<<ComboboxSelected>>', self.actualizartreeEOP1)
            #self.cb_periodo.current(0)
            self.cb_area_op.place(x=520, y=5)

            #Boton para agregar eva        
            self.btn_agPago = tk.Button(panel_principal, text="Registrar pago", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.registrarOP)
            self.btn_agPago.place(x=800, y=290)

            #Boton para agregar eva        
            self.btn_savePago = tk.Button(panel_principal, text="Salvar pagos", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.salvarOP)
            self.btn_savePago.place(x=800, y=350)

            #Boton aprobar evaluaciones        
            self.btn_signEmpOP = tk.Button(panel_principal, text="Revisar otros pagos", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.signOP)
            self.btn_signEmpOP.place(x=800, y=410)

            #label empleado
            self.lb_sempleado_op = tk.Label(panel_principal, text='Empleado seleccionado', justify='right', bg=COLOR_CUERPO_PRINCIPAL, font=('Times', 12), width=25)
            self.lb_sempleado_op.place(x=800, y=100)       
            
            #label tipo pago
            self.lb_tp_op = tk.Label(panel_principal, text='Mes 1', justify='right', bg=COLOR_CUERPO_PRINCIPAL, font=('Times', 12), width=10)
            self.lb_tp_op.place(x=784, y=150)
            #label mes2
            self.lb_monto_op = tk.Label(panel_principal, text='Mes 2', justify='right', bg=COLOR_CUERPO_PRINCIPAL, font=('Times', 12), width=10)
            self.lb_monto_op.place(x=784, y=200)
            

            #ComboEva x mes
            self.cb_tp_op = ttk.Combobox(panel_principal, postcommand=self.cargarTP, width=10)
            #self.cb_periodo.current(0)
            self.cb_tp_op.place(x=870, y=150)

            self.tx_monto_op = ttk.Entry(panel_principal, font=('Times', 14), width=10)
            #self.cb_periodo.current(0)
            self.tx_monto_op.place(x=870, y=200)
      
            

            #Treeview        
            columns = ('numero', 'nombreap', 'ci', 'area','opagos')
            self.treeEOP = ttk.Treeview(panel_principal, height=16, columns=columns, show='headings')
            self.style = ttk.Style(self.treeEOP)
            self.style.configure('Treeview',rowheight=30)
            self.treeEOP.column('numero',width=80)
            self.treeEOP.column('nombreap',width=200)
            self.treeEOP.column('ci',width=110)
            self.treeEOP.column('area',width=200)
            self.treeEOP.column('opagos',width=100)

            self.treeEOP.heading(column='numero', text='No.')
            self.treeEOP.heading(column='nombreap', text='Nombre y apellidos')
            self.treeEOP.heading(column='ci', text='CI')
            self.treeEOP.heading(column='area', text='Área')
            self.treeEOP.heading(column='opagos',text='Otros pagos')
            self.treeEOP.grid(row=1,column=0, columnspan=5,ipadx=5,padx=5,pady=5)
            self.treeEOP.bind('<<TreeviewSelect>>',self.selectEmp)
            self.actualizartreeEOP()  
            self.cargarTP()  
            self.cargarDpto()   
        else:
            messagebox.showinfo('Notificación','Debe registrar un período de evaluación')


    #Definiendo tree view de periodo
    def registrarOP(self): 
        pass       
        # curItem = self.treeEOP.focus()
        # if curItem:   
        #     self.treeEOP.set(curItem, column='mes1', value=self.cb_tp_op.get())   
        # else:            
        #     messagebox.showinfo('Información','Debe seleccionar un trabajador')

        
    def actualizartreeEOP(self):
        self.treeEOP.delete(*self.treeEOP.get_children())         
        queryEmpL=''
        if self.tx_empleado_op.get() != '' and self.cb_area_op.get() == '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id where x.nombreap like '%"+self.tx_empleado_op.get().upper()+"% ORDER BY a.id'"
        elif self.cb_area_op.get() != '' and self.tx_empleado_op.get() == '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id where a.area = '"+self.cb_area_op.get()+" ORDER BY a.id'"
        elif self.tx_empleado_op.get() != '' and self.cb_area_op.get() != '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id where x.nombreap like '%"+self.tx_empleado_op.get().upper()+"%' and a.area = '"+self.cb_area_op.get()+" ORDER BY a.id'"
        else:
            queryEmpL='SELECT x.id,x.nombreap,x.ci,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id  ORDER BY a.id'
        
         
        self.cursorLoc.execute(queryEmpL)

        slistEmp = self.cursorLoc.fetchall()            
        for row in slistEmp:
            self.treeEOP.insert('','end',values=("'"+row[0]+"'",row[1],row[2],row[3])) 


    def actualizartreeEOP1(self,event):
        self.actualizartreeEOP()

    def selectEmp(self,event):
        curItem = self.treeEOP.focus()
        selectedItem=self.treeEOP.item(curItem)
        cadena=str(selectedItem['values'][0])+" "+str(selectedItem['values'][1])
        if len(cadena) <26:
            self.lb_sempleado_op['text']=cadena
        else:
            self.lb_sempleado_op['text']=cadena[0:20]+"..."
    

    def cargarTP(self):
        options=[]         
        queryP='SELECT x.* FROM postgres.public.tipo_pago x order by id asc'
        self.cursorLoc.execute(queryP)
        slistTP=self.cursorLoc.fetchall()
        for row in slistTP:
            options.append(str(row[0])+'-'+row[1])
        
        self.cb_tp_op['values']=options
    
    def cargarDpto(self):
        options=[]         
        queryP='SELECT x.* FROM postgres.public.area x order by area asc'
        self.cursorLoc.execute(queryP)
        slistArea=self.cursorLoc.fetchall()
        for row in slistArea:
            options.append(row[1])
        
        self.cb_area_op['values']=options

    def getPeriodo(self):         
        queryP='SELECT p.* FROM postgres.public.utilidades_periodo_incluye x INNER JOIN postgres.public.periodo AS p ON x.upincluye_periodo_id = p.id order by p.id asc'
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()

    def obtenerPerMes(self,mes):
        queryP="SELECT * FROM postgres.public.periodo where mes='"+str(mes)+"'"
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()  

    
    
    def salvarOP(self):
        pass
        # try:
        #     for parent in self.treeEOP.get_children():  
        #         listTree=self.treeEOP.item(parent)["values"]
        #         slistP=self.getPeriodo()
        #         evames1 = self.obtenerEvaCond(listTree[0],str(slistP[0][0]))
        #         evames2 = self.obtenerEvaCond(listTree[0],str(slistP[1][0]))
        #         evames3 = self.obtenerEvaCond(listTree[0],str(slistP[2][0]))
        #         if evames1 is not None:
        #             queryInEvaM1 = "UPDATE postgres.public.evaluacion SET evaluacion_tipoevaluacion_id = "+str(self.obtenerTipoEva(listTree[4])[0])+" \
        #                 WHERE evaluacion_empleado_id = "+listTree[0]+" and evaluacion_perio_id = "+str(slistP[0][0])
        #         else:
        #             queryInEvaM1="INSERT INTO postgres.public.evaluacion (evaluacion_empleado_id,evaluacion_tipoevaluacion_id,evaluacion_perio_id) \
        #                     VALUES ("+listTree[0]+","+str(self.obtenerTipoEva(listTree[4])[0][0])+","+str(slistP[0][0])+")"
        #         if evames2 is not None:
        #             queryInEvaM2 = "UPDATE postgres.public.evaluacion SET evaluacion_tipoevaluacion_id = "+str(self.obtenerTipoEva(listTree[5])[0])+" \
        #                 WHERE evaluacion_empleado_id = "+listTree[0]+" and evaluacion_perio_id = "+str(slistP[1][0])
        #         else:
        #             queryInEvaM2="INSERT INTO postgres.public.evaluacion (evaluacion_empleado_id,evaluacion_tipoevaluacion_id,evaluacion_perio_id) \
        #                     VALUES ("+listTree[0]+","+str(self.obtenerTipoEva(listTree[5])[0][0])+","+str(slistP[1][0])+")"
        #         if evames3 is not None:
        #             queryInEvaM3 = "UPDATE postgres.public.evaluacion SET evaluacion_tipoevaluacion_id = "+str(self.obtenerTipoEva(listTree[6])[0])+" \
        #                 WHERE evaluacion_empleado_id = "+listTree[0]+" and evaluacion_perio_id = "+str(slistP[2][0])
        #         else:            
        #             queryInEvaM3="INSERT INTO postgres.public.evaluacion (evaluacion_empleado_id,evaluacion_tipoevaluacion_id,evaluacion_perio_id) \
        #                     VALUES ("+listTree[0]+","+str(self.obtenerTipoEva(listTree[6])[0][0])+","+str(slistP[2][0])+")"
        #         self.cursorLoc.execute(queryInEvaM1)
        #         self.cursorLoc.execute(queryInEvaM2)
        #         self.cursorLoc.execute(queryInEvaM3)
        #         self.connLoc.commit()
        #     messagebox.showinfo('Confirmación','Las evaluaciones se registraron satisfactoriamente')
        # except Exception as error:
        #         messagebox.showerror("Error",error)    

    

    def getDepartamento(self,idemp):         
        queryP="SELECT a.area  FROM postgres.public.empleado emp INNER JOIN postgres.public.area AS a ON emp.empleado_area_id  = a.id where emp.id = "+str(idemp)
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchone()[0]

    def signOP(self): 
               
        path = "file/evaluacion.xlsx"
        self.limpiarExcel(3,path) 
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        row = 3
        sheet['C2']=self.getPeriodo()[0][1]
        sheet['D2']=self.getPeriodo()[1][1]
        sheet['E2']=self.getPeriodo()[2][1]
        for parent in self.treeEOP.get_children():
            values=self.treeEOP.item(parent)["values"]
            #Insertar empleados
            sheet['A'+str(row)]=self.getDepartamento(values[0])
            sheet['B'+str(row)]=values[1]
            sheet['C'+str(row)]=values[4]
            sheet['D'+str(row)]=values[5]
            sheet['E'+str(row)]=values[6]

            row+=1
        wb.save(path)
        self.convert_xlsx_to_pdf(path)

    def convert_xlsx_to_pdf(self,xlsx_file):
        try:
            subprocess.run(["libreoffice24.2", "--headless", "--convert-to", "pdf", xlsx_file])
            separador = os.path.sep
            dir_actual = os.path.dirname(os.path.abspath(__file__))
            dir = separador.join(dir_actual.split(separador)[:-1])
            #dirfile = separador.join(xlsx_file.split(separador))
            command =  ['open', dir+separador+'evaluacion.pdf']
            subprocess.run(command,shell=False)
            print("Done!")

        except Exception as e:
            print("Error:", e)


    

    

    def limpiarExcel(self,fila,url):         
        path = url
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        sheet.delete_rows(fila, sheet.max_row-1)        
        wb.save(path)

    
