import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR,CONN_LOC,CURSOR_LOC,CONN_ZUN,CURSOR_ZUN
import openpyxl
import os
import subprocess
import tkinter.font as tkfont
from openpyxl.styles import Font, colors, fills, Alignment, PatternFill, NamedStyle



class FormularioCalcImpDesign():
    
    def __init__(self, panel_principal):   
       
        #Definiendo variables
        # Variables de conexion
        self.connLoc = CONN_LOC
        self.cursorLoc = CURSOR_LOC
        self.conn = CONN_ZUN
        self.cursor = CURSOR_ZUN
        if self.getPeriodo():
            
            # Definiendo controles de seleccion
            self.empSelec = ''
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

            #Boton para Registrar pago        
            self.btn_agPago = tk.Button(panel_principal, text="Informe/utilidades", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.registrarOP)
            self.btn_agPago.place(x=835, y=150)     

            #Boton para Eliminar pago        
            self.btn_agPago = tk.Button(panel_principal, text="Informe. x Depart.", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.eliminarOP)
            self.btn_agPago.place(x=835, y=200)             


            #Boton Revisar otros pagos        
            self.btn_signEmpOP = tk.Button(panel_principal, text="Exportar dbf", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.signOP)
            self.btn_signEmpOP.place(x=835, y=250)
            

            #Periodo del pago
            self.lx_PPlabel = tk.Label(panel_principal, text="Mes de pago:", justify='right', bg=COLOR_CUERPO_PRINCIPAL, font=('Times', 12))
            self.lx_PPlabel.place(x=850, y=80) 
            
            self.cb_periodo_op = ttk.Combobox(panel_principal, postcommand=self.getMesesA, width=12)
            #self.cb_periodo.current(0)
            self.cb_periodo_op.place(x=850, y=100) 
      
            

            #Treeview        
            columns = ('numero', 'ci', 'nombreap', 'devutil','segsoc', 'imping', 'descrm', 'neto')
            self.treeEOP = ttk.Treeview(panel_principal, height=16, columns=columns, show='headings')
            self.style = ttk.Style(self.treeEOP)
            self.style.configure('Treeview',rowheight=30)
            self.treeEOP.column('numero',width=80)
            self.treeEOP.column('ci',width=110)
            self.treeEOP.column('nombreap',width=200)            
            self.treeEOP.column('devutil',width=100)
            self.treeEOP.column('segsoc',width=80)
            self.treeEOP.column('imping',width=80)
            self.treeEOP.column('descrm',width=80)
            self.treeEOP.column('neto',width=80)

            self.treeEOP.heading(column='numero', text='No.')
            self.treeEOP.heading(column='ci', text='CI')
            self.treeEOP.heading(column='nombreap', text='Nombre y apellidos')            
            self.treeEOP.heading(column='devutil', text='Dev/Utili')
            self.treeEOP.heading(column='segsoc',text='Seg/Soc')
            self.treeEOP.heading(column='imping',text='Imp/Ing')
            self.treeEOP.heading(column='descrm',text='Desc/RM')
            self.treeEOP.heading(column='neto',text='Neto')
            
            
            self.treeEOP.grid(row=1,column=0, columnspan=5,ipadx=5,padx=5,pady=5)
            #self.actualizartreeEOP()  
              
            self.cargarDpto()   
        else:
            messagebox.showinfo('Notificación','Debe registrar un período de evaluación')

    #Listado de otros pagos
    

    #Cargar combo de periodo
    def cargarPeriodoOP(self): 
        slistp = self.getPeriodo() 
        options = []       
        for row in slistp:
            options.append(str(row[0])+"-"+str(row[1]))  

        self.cb_periodo_op['value']=options

    #Definiendo tree view de periodo
    def registrarOP(self):       
        if self.empSelec:
            idTPSelected = self.cb_tp_op.get().split('-')[0]
            idPeriodo = self.cb_periodo_op.get().split('-')[0]
            if self.tx_monto_op.get() != '' and  idPeriodo!= '' and  idTPSelected != '':
                selectedItem=self.treeEOP.item(self.empSelec)
                queryIOP = "INSERT INTO postgres.public.opago (tpago_id, monto, opago_periodo_id, opago_empleado_id) \
                    VALUES("+str(idTPSelected)+","+self.tx_monto_op.get()+","+str(idPeriodo)+","+str(selectedItem['values'][0])+")"
                self.cursorLoc.execute(queryIOP)
                self.connLoc.commit()            
                #self.treeEOP.set(self.empSelec, column='opagos', value=self.cb_tp_op.get())  
                messagebox.showinfo('Confirmación','La información del pago se registró satisfactoriamente') 
            else:
                messagebox.showinfo('Campos vacíos','Existen campos vacíos, debe completarlos')
        else:            
            messagebox.showinfo('Información','Debe seleccionar un trabajador')
        self.actualizartreeEOP()

    def eliminarOP(self):
        if self.empSelec:
            idTPSelected = self.cb_tp_op.get().split('-')[0]
            idPeriodo = self.cb_periodo_op.get().split('-')[0]
            selectedItem=self.treeEOP.item(self.empSelec)
            cantOPEmpbefore = len(self.listOP(selectedItem['values'][0]))
            if self.tx_monto_op.get() != '' and  idPeriodo!= '' and  idTPSelected != '':                
                queryEOP = "DELETE FROM postgres.public.opago WHERE tpago_id = "+str(idTPSelected)+"\
                        AND monto = "+self.tx_monto_op.get()+" AND opago_periodo_id = "+str(idPeriodo)+" AND opago_empleado_id = "+str(selectedItem['values'][0])
                self.cursorLoc.execute(queryEOP)
                self.connLoc.commit() 
                cantP = len(self.listOP(selectedItem['values'][0]))
                if cantOPEmpbefore == cantP:
                    messagebox.showinfo('Sin acción','No existen registros para la información suministrada')
                    #self.treeEOP.set(self.empSelec, column='opagos', value=self.cb_tp_op.get())  
                else:
                    messagebox.showinfo('Confirmación','La información se eliminó correctamente') 
            else:
                messagebox.showinfo('Campos vacíos','Existen campos vacíos, debe completarlos')
        else:            
            messagebox.showinfo('Información','Debe seleccionar un trabajador')
        self.actualizartreeEOP()


        
    def actualizartreeEOP(self):
        self.treeEOP.delete(*self.treeEOP.get_children())         
        queryEmpL=''
        if self.tx_empleado_op.get() != '' and self.cb_area_op.get() == '':
            queryEmpL="SELECT x.* FROM postgres.public.utilidades_printhist x where x.nombreap like '%"+self.tx_empleado_op.get().upper()+"%' ORDER BY x.area ASC"
        elif self.cb_area_op.get() != '' and self.tx_empleado_op.get() == '':
            queryEmpL="SELECT x.* FROM postgres.public.utilidades_printhist x where x.area = '"+self.cb_area_op.get()+"' ORDER BY x.area ASC"
        elif self.tx_empleado_op.get() != '' and self.cb_area_op.get() != '':
            queryEmpL="SELECT x.* FROM postgres.public.utilidades_printhist x where x.nombreap like '%"+self.tx_empleado_op.get().upper()+"%' and x.area = '"+self.cb_area_op.get()+"' ORDER BY a.id ASC"
        else:
            queryEmpL='SELECT x.* FROM postgres.public.utilidades_printhist x ORDER BY a.id'
        
         
        self.cursorLoc.execute(queryEmpL)

        slistEmp = self.cursorLoc.fetchall()            
        for emp in slistEmp:
            self.treeEOP.insert('','end',values=("'"+emp[0]+"'",emp[3],emp[4],emp[5],emp[6],emp[7],emp[8]))


    def actualizartreeEOP1(self,event):
        self.actualizartreeEOP()  
    

    
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

    def getUtiliDist(self):
        queryUD="SELECT x.* FROM postgres.public.utilidades_distribucion x"
        self.cursorLoc.execute(queryUD)
        result = self.cursorLoc.fetchone()
        return result 


    def getMesesA(self):
        if self.getPeriodo()[2][0] == 'Diciembre':
            queryP="SELECT id_peri, nombre FROM ZUNpr.dbo.n_periodo where fecha_inicio like '%"+str(self.getUtiliDist()[3]+1)+"%'"
        else:
            queryP="SELECT id_peri, nombre FROM ZUNpr.dbo.n_periodo where fecha_inicio like '%"+str(self.getUtiliDist()[3])+"%'"

        self.cursor.execute(queryP)
        listmeseA = self.cursor.fetchall() 
        options=[]
        for mes in listmeseA:
            if mes['id_peri'] > self.getPeriodo()[2][0]:
                options.append((str(mes['id_peri'])+'-'+str(mes['nombre'])))
        self.cb_periodo_op['values']=options

    def obtenerPerMes(self,mes):
        queryP="SELECT * FROM postgres.public.periodo where mes='"+str(mes)+"'"
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()      

    def getDepartamento(self,idemp):         
        queryP="SELECT a.area  FROM postgres.public.empleado emp INNER JOIN postgres.public.area AS a ON emp.empleado_area_id  = a.id where emp.id = "+str(idemp)
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchone()[0]


    #Mostrar reportes de otros pagos    
    def signOP(self): 
               
        path = "file/list_opagos.xlsx"
        self.limpiarExcel(4,path) 
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        consec = 1
        row = 4
        text_format = Font(
            bold = True,
            name = 'Calibri',
            size = '-1',
            color = colors.BLACK )   
        sheet['G3']=self.getPeriodo()[0][1].upper()
        sheet['G3'].font += text_format 
        sheet['H3']=self.getPeriodo()[1][1].upper()
        sheet['H3'].font += text_format
        sheet['I3']=self.getPeriodo()[2][1].upper()
        sheet['I3'].font += text_format
        
        
        for parent in self.treeEOP.get_children():
            values=self.treeEOP.item(parent)["values"]
            countmerge = len(self.listOP(values[0]))
            sheet["A"+str(row)].alignment = Alignment(vertical='top')
            sheet["B"+str(row)].alignment = Alignment(vertical='top')
            sheet["C"+str(row)].alignment = Alignment(vertical='top')
            sheet["D"+str(row)].alignment = Alignment(vertical='top')
            sheet["E"+str(row)].alignment = Alignment(vertical='top')         
            
            
            if len(self.listOP(values[0])) > 0:
                
                if countmerge > 1:
                    sheet.merge_cells("A"+str(row)+":A"+str(row+countmerge-1))                     
                    sheet.merge_cells("B"+str(row)+":B"+str(row+countmerge-1))                    
                    sheet.merge_cells("C"+str(row)+":C"+str(row+countmerge-1))                    
                    sheet.merge_cells("D"+str(row)+":D"+str(row+countmerge-1))                    
                    sheet.merge_cells("E"+str(row)+":E"+str(row+countmerge-1))                    

                sheet['A'+str(row)]=consec
                sheet['B'+str(row)]=values[0]
                sheet['C'+str(row)]=values[1]
                sheet['D'+str(row)]=values[2]
                sheet['E'+str(row)]=values[3]
                if self.listOP(values[0],self.getPeriodo()[0][0]) is not None:
                    listop = self.listOP(values[0],self.getPeriodo()[0][0])
                    for p in listop:
                        sheet['F'+str(row)]=p[5]   
                        sheet["F"+str(row)].alignment = Alignment(vertical='top', wrapText=True)                    
                        sheet['G'+str(row)]=p[4]
                        sheet["G"+str(row)].alignment = Alignment(vertical='top', horizontal = 'center')
                        row+=1
                if self.listOP(values[0],self.getPeriodo()[1][0]) is not None:
                    listop = self.listOP(values[0],self.getPeriodo()[1][0])
                    for p in listop:
                        sheet['F'+str(row)]=p[5] 
                        sheet["F"+str(row)].alignment = Alignment(vertical='top', wrapText=True)                       
                        sheet['H'+str(row)]=p[4]
                        sheet["H"+str(row)].alignment = Alignment(vertical='top', horizontal = 'center')
                        row+=1
                if self.listOP(values[0],self.getPeriodo()[2][0]) is not None:
                    listop = self.listOP(values[0],self.getPeriodo()[2][0])
                    for p in listop:
                        sheet['F'+str(row)]=p[5]
                        sheet["F"+str(row)].alignment = Alignment(vertical='top', wrapText=True)
                        sheet['I'+str(row)]=p[4]
                        sheet["I"+str(row)].alignment = Alignment(vertical='top', horizontal = 'center')
                        row+=1
                
                consec+=1
            #Insertar empleados
            
            
        wb.save(path)
        self.convert_xlsx_to_pdf(path,'list_opagos')
        self.unmergexlsop()

    def convert_xlsx_to_pdf(self,xlsx_file,nombreA=''):
        try:
            subprocess.run(["libreoffice24.2", "--headless", "--convert-to", "pdf", xlsx_file])
            separador = os.path.sep
            dir_actual = os.path.dirname(os.path.abspath(__file__))
            dir = separador.join(dir_actual.split(separador)[:-1])
            #dirfile = separador.join(xlsx_file.split(separador))
            command =  ['open', dir+separador+nombreA+'.pdf']
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

    
