import tkinter as tk
from tkinter import *
import pymssql
import psycopg2
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR
from PIL import Image, ImageTk
import openpyxl


class FormularioCargarEDesign():

    def __init__(self, panel_principal):   
       
        #Definiendo variables
        # Definiendo controles         
        self.im_checked = ImageTk.PhotoImage(Image.open("imagenes/checked.png").resize((15,15)))
        self.im_unchecked = ImageTk.PhotoImage(Image.open("imagenes/unchecked.png").resize((15,15)))
        #Conexion
        self.connZun = pymssql.connect(
            server='10.105.213.6',
            user='userutil',
            password='1234',
            database='ZUNpr',
            as_dict=True)
        self.cursorZun = self.connZun.cursor()

        self.connLoc = psycopg2.connect(
            host="localhost",
            database="postgres",
            user="postgres",
            password="proyecto") 
        self.cursorLoc = self.connLoc.cursor() 
        if self.getPeriodo():
            #type empleado
            self.tx_empleado = ttk.Entry(panel_principal, font=(
                'Times', 14), width=10)
            self.tx_empleado.grid(row=0,column=0,padx=5,pady=5,ipadx=40)

            #Boton para cargar empleados            
            self.btn_cargar = tk.Button(panel_principal, text="Cargar empleados", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.cargartreeE)
            self.btn_cargar.place(x=840, y=150)

            #Boton para guardar empleados            
            self.btn_cargar = tk.Button(panel_principal, text="Registrar destajos", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.guardarListado)
            self.btn_cargar.place(x=840, y=200)

            #Boton para mostrar empleados            
            self.btn_cargar = tk.Button(panel_principal, text="Exportar listado", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.mostrarListado)
            self.btn_cargar.place(x=840, y=250)

            #Boton para buscar empleados
            
            self.btn_bempleados = tk.Button(panel_principal, text="Buscar", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.actualizarTreeE)
            #self.btn_bempleados.grid(row=0,column=2,padx=5,pady=5)
            self.btn_bempleados.place(x=255, y=5)

            self.tx_area = tk.Label(panel_principal, font=('Times', 14), width=20, bg=COLOR_CUERPO_PRINCIPAL, text='Departamento:')
            self.tx_area.place(x=350, y=5)

            #Combo departamento
            self.cb_area= ttk.Combobox(panel_principal, width=30)
            self.cb_area.bind('<<ComboboxSelected>>', self.actualizarTreeE1)
            #self.cb_periodo.current(0)
            self.cb_area.place(x=520, y=5)

            #Treeview
                    
            
            self.treeE = ttk.Treeview(panel_principal,height=16,columns=(1,2,3,4,5))
            self.style = ttk.Style(self.treeE)
            self.style.configure('Treeview',rowheight=30)

            self.treeE.tag_configure('checked', image=self.im_checked)
            self.treeE.tag_configure('unchecked', image=self.im_unchecked)
            
            self.treeE.column('1',width=80)
            self.treeE.column('2',width=200)
            self.treeE.column('3',width=130)
            self.treeE.column('4',width=100)
            self.treeE.column('5',width=100)
            

            self.treeE.heading('#0', text='Destajo',anchor = CENTER)
            self.treeE.heading('#1', text='No.',anchor = CENTER)
            self.treeE.heading('#2', text='Nombre y apellidos',anchor = CENTER)
            self.treeE.heading('#3', text='CI',anchor = CENTER)
            self.treeE.heading('#4', text='Escala',anchor = CENTER)
            self.treeE.heading('#5', text='TarifaH',anchor = CENTER)
            
            self.treeE.grid(row=1,column=0, columnspan=5,ipadx=5,padx=5,pady=5)
            #self.treeE.bind('<Double 1>', self.getrow)
            self.treeE.bind('<Button 1>', self.toggleCheck)

            self.actualizarTreeE()
            self.cargarDpto()
        else:
            messagebox.showinfo('Notificación','Debe registrar un período de evaluación')
        

    #Obtener row seleccionado
    # def getrow(self):
    #     item = self.treeE.item(self.treeE.focus())
    #     l1.set(item['values'][0])

    def toggleCheck(self,event):
        rowid=self.treeE.identify_row(event.y)
        tag = self.treeE.item(rowid,"tags")[0]
        tags = list(self.treeE.item(rowid,"tags"))
        tags.remove(tag)
        self.treeE.item(rowid,tags=tags)
        if tag == 'checked':
           self.treeE.item(rowid,tags='unchecked') 
        else:
           self.treeE.item(rowid,tags='checked') 

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

    #Verificar Emp en los 3 meses del periode devolviendo su tarifahorario 
    def verificarEmpTrim(self,nointerno):
        tarifaH=tarifatemp=''
        
        queryEmpT="SELECT x.tarifa_horaria FROM ZUNpr.dbo.h_empleado x where x.no_interno='"+nointerno+"' and x.id_peri="+str(self.getPeriodo()[0][0])        
        self.cursorZun.execute(queryEmpT) 
        tarifatemp=self.cursorZun.fetchone()      
        if tarifatemp is not None: 
            tarifaH=tarifatemp['tarifa_horaria']   
        queryEmpT="SELECT x.tarifa_horaria FROM ZUNpr.dbo.h_empleado x where x.no_interno='"+nointerno+"' and x.id_peri="+str(self.getPeriodo()[1][0])
        self.cursorZun.execute(queryEmpT)
        tarifatemp=self.cursorZun.fetchone()
        if tarifatemp is not None:
            tarifaH=tarifatemp['tarifa_horaria']
        queryEmpT="SELECT x.tarifa_horaria FROM ZUNpr.dbo.h_empleado x where x.no_interno='"+nointerno+"' and x.id_peri="+str(self.getPeriodo()[2][0])
        self.cursorZun.execute(queryEmpT)
        tarifatemp=self.cursorZun.fetchone() 
        if tarifatemp is not None:
            tarifaH=tarifatemp['tarifa_horaria']  

        
        return tarifaH

    def limpiarEmpArea(self):
        queryDemp = 'DELETE FROM postgres.public.area'
        self.cursorLoc.execute(queryDemp)
        self.connLoc.commit()


    #Cargar tree de empleado desde el zun
    def cargartreeE(self):
        self.treeE.delete(*self.treeE.get_children())
        self.limpiarEmpArea()
        try:            
            queryEmp="SELECT e1.no_interno,e1.nombre,e1.apell1,e1.apell2,e1.no_expediente,e1.cargo,nuo.descripcion,nuo.id_uorg,ge.iden_grupo_sal \
                FROM ZUNpr.dbo.p_empleado AS e1 INNER JOIN ZUNpr.dbo.nomina_sal AS noms ON e1.no_interno = noms.no_interno \
                INNER JOIN ZUNpr.dbo.n_grupo_escala AS ge ON e1.grupo_escala = ge.id_grupo_sal  \
                INNER JOIN ZUNpr.dbo.n_unidad_org AS nuo ON e1.unidad_org = nuo.id_uorg \
                WHERE noms.id_periodo >= "+str(self.getPeriodo()[0][0])+" AND noms.id_periodo <= "+str(self.getPeriodo()[2][0])+" \
                GROUP BY e1.no_interno,e1.nombre,e1.apell1,e1.apell2,e1.no_expediente,e1.cargo,nuo.descripcion,nuo.id_uorg,ge.iden_grupo_sal \
                ORDER BY nuo.id_uorg ASC"
            departamento = ''                      
            #importar empleados            
            self.cursorZun.execute(queryEmp)
            slistEmp = self.cursorZun.fetchall()
            for row in slistEmp:    
                
                tarifaH=self.verificarEmpTrim(row['no_interno'])
                if tarifaH != '':
                    if departamento != row['descripcion']:
                        departamento = row['descripcion']
                        insertArea = "INSERT INTO postgres.public.area (id, area) VALUES("+str(row['id_uorg'])+",'"+row['descripcion']+"')"
                        self.cursorLoc.execute(insertArea)
                        self.connLoc.commit()
                    queryInsertEmp="INSERT INTO postgres.public.empleado \
                    (id, nombreap, ci, escalas, thoraria, destajo, empleado_area_id)\
                    VALUES('"+row['no_interno']+"', '"+row['nombre']+" "+row['apell1']+" "+row['apell2']+"', '"+row['no_expediente']+"', '"+row['iden_grupo_sal'].rstrip()+"', "+str(tarifaH)+", False, "+str(row['id_uorg'])+")"
                    #print(queryInsertEmp)                    
                    self.cursorLoc.execute(queryInsertEmp)
                    self.connLoc.commit() 

                    #self.treeE.insert('','end',values=("'"+row['no_interno']+"'",row['nombre']+" "+row['apell1']+" "+row['apell2'],"'"+row['no_expediente']+"'",row['iden_grupo_sal'].rstrip(),tarifaH),tags='unchecked')
                else:
                    continue
            self.cargarDpto()
            self.actualizarTreeE()
                    
                
        except Exception as error:
            messagebox.showerror("Error",error)

        
    def actualizarTreeE(self):
        self.treeE.delete(*self.treeE.get_children())         
        queryEmpL=''
        if self.tx_empleado.get() != '' and self.cb_area.get() == '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,x.escalas,x.thoraria,x.destajo,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id where x.nombreap like '%"+self.tx_empleado.get().upper()+"%' ORDER BY a.id ASC"
        elif self.cb_area.get() != '' and self.tx_empleado.get() == '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,x.escalas,x.thoraria,x.destajo,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id where a.area = '"+self.cb_area.get()+"' ORDER BY a.id ASC"
        elif self.tx_empleado.get() != '' and self.cb_area.get() != '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,x.escalas,x.thoraria,x.destajo,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id where x.nombreap like '%"+self.tx_empleado.get().upper()+"%' and a.area = '"+self.cb_area.get()+"' ORDER BY a.id ASC"
        else:
            queryEmpL='SELECT x.id,x.nombreap,x.ci,x.escalas,x.thoraria,x.destajo,a.area FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id ORDER BY a.id ASC'
               
        self.cursorLoc.execute(queryEmpL)

        slistEmp = self.cursorLoc.fetchall()       
        for row in slistEmp:
            if row[5]==False:
                self.treeE.insert('','end',values=("'"+row[0]+"'",row[1],row[2],row[3],row[4],row[6]),tags='unchecked')
            else:
                self.treeE.insert('','end',values=("'"+row[0]+"'",row[1],row[2],row[3],row[4],row[6]),tags='checked')

    def guardarListado(self):
        try:
            for parent in self.treeE.get_children():
                #Insertar empleados
                tag=self.treeE.item(parent)["tags"]
                values=self.treeE.item(parent)["values"]
                if tag[0] == 'checked':
                    querySETDestajo = "UPDATE postgres.public.empleado SET destajo=true WHERE id="+str(values[0])
                    self.cursorLoc.execute(querySETDestajo)
                    self.connLoc.commit()
                else:
                    querySETDestajo = "UPDATE postgres.public.empleado SET destajo=false WHERE id="+str(values[0])
                    #print(querySETDestajo)
                    self.cursorLoc.execute(querySETDestajo)
                    self.connLoc.commit()
            messagebox.showinfo('Confirmación','Los destajos se registraron satisfactoriamente')
        except Exception as error:
                messagebox.showerror("Error",error)
    
    def getPeriodo(self):         
        queryP='SELECT p.* FROM postgres.public.utilidades_periodo_incluye x INNER JOIN postgres.public.periodo AS p ON x.upincluye_periodo_id = p.id order by p.id asc'
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()

    def mostrarListado(self):         
        path = "file/Listado de trabajadores1.xlsx"

        wb = openpyxl.load_workbook(path)

        sheet = wb.active
        row = 5
        for parent in self.treeE.get_children():
            #Insertar empleados
            tag=self.treeE.item(parent)["tags"]
            values=self.treeE.item(parent)["values"]
            sheet['A'+str(row)]=self.getDepartamento(values[0])
            sheet['B'+str(row)]=values[0]
            sheet['C'+str(row)]=values[1]
            sheet['D'+str(row)]=values[2]
            sheet['E'+str(row)]=values[3]
            sheet['F'+str(row)]=values[4]
            if tag[0] == 'checked':
                sheet['G'+str(row)]='Si'
            else:
                sheet['G'+str(row)]='No'

            row+=1


        wb.save(path)

    def getDepartamento(self,idemp):         
        queryP="SELECT a.area  FROM postgres.public.empleado emp INNER JOIN postgres.public.area AS a ON emp.empleado_area_id  = a.id where emp.id = "+str(idemp)
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchone()[0]
    