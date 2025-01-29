import tkinter as tk
from tkinter import *
import pymssql
import psycopg2
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR
from datetime import datetime



class FormularioRegistroPDesign():

    def __init__(self, panel_principal):   
       
        #Definiendo variables de conexion
        self.conn = pymssql.connect(
            server='10.105.213.6',
            user='userutil',
            password='1234',
            database='ZUNpr',
            as_dict=True
            )
        self.cursor = self.conn.cursor()
        self.connLoc = psycopg2.connect(host="localhost", database="postgres", user="postgres", password="proyecto")
        self.cursorLoc = self.connLoc.cursor()
        #Definiendo variables
        self.var_periodo = StringVar()
        self.anno_trim = StringVar()
        self.periodoregistrado=[]
        self.registrado = False
        # Definiendo controles 
        
        #comobo carga periodo del zun
        self.cbx_label = tk.Label(panel_principal, text="Mes", bg=COLOR_CUERPO_PRINCIPAL)
        self.cbx_label.grid(row=1,column=0,padx=5,pady=10)
        
        self.cb_periodo = ttk.Combobox(panel_principal,textvariable=self.var_periodo, postcommand=self.cargarcombo)
        #self.cb_periodo.current(0)
        self.cb_periodo.grid(row=1,column=1,padx=5,pady=10)          

        

        #nombre del periodo
        self.tx_label = tk.Label(panel_principal, text="Nombre del período", bg=COLOR_CUERPO_PRINCIPAL)
        self.tx_label.grid(row=0,column=0,padx=5,pady=10)
        
        self.tx_trimestre_name = ttk.Combobox(panel_principal, font=('Times', 14), values=('1er Trimestre','2do Trimestre','3er Trimestre','4to Trimestre'))
        self.tx_trimestre_name.grid(row=0,column=1,padx=5,pady=10)

        #Año del periodo
        self.tx_labelA = tk.Label(panel_principal, text="Año", bg=COLOR_CUERPO_PRINCIPAL)
        self.tx_labelA.grid(row=0,column=2,padx=5,pady=10)
                
        self.annot = str(self.var_periodo)[:4]
        #range(10)
        optionsAno=list(range((datetime.now().year-10),(datetime.now().year+1)))
        self.tx_anno = ttk.Combobox(panel_principal, font=('Times', 14), values=optionsAno)
        self.tx_anno.grid(row=0,column=3,padx=5,pady=10)

        #Boton para agregar periodo al treewiew
        self.btn_registro_periodo = tk.Button(panel_principal, text="Registrar", font=(
            'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, padx=5, command=self.addperiodo)
        self.btn_registro_periodo.grid(row=1,column=3,padx=5,pady=10)
        

        #Boton para crear Trimestre
        self.btn_save = tk.Button(panel_principal, text="Guargar periodo", font=(
            'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, padx=15, command=self.save)
        self.btn_save.grid(row=3,columnspan=5,padx=10,pady=10)
        #Boton para reiniciar Trimestre
        self.btn_save = tk.Button(panel_principal, text="Reiniciar periodo", font=(
            'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, padx=15, command=self.reiniciarPeriodo)
        self.btn_save.grid(row=3,columnspan=3,padx=10,pady=10)


        #Definiendo tree view de periodo

        self.tree = ttk.Treeview(panel_principal,
                                 show='headings')
        self.tree['columns'] = ('id', 'periodo', 'fechainicio', 'fechafin', 'orden')
        self.tree.column('id',width=80)
        self.tree.column('periodo',width=100)        
        self.tree.column('fechainicio',width=150)
        self.tree.column('fechafin',width=150)
        self.tree.column('orden',width=80)

        self.tree.heading('id', text='Id')
        self.tree.heading('periodo', text='Período')        
        self.tree.heading('fechainicio', text='Fecha inicio')
        self.tree.heading('fechafin', text='Fecha fin')
        self.tree.heading('orden', text='Orden')
        self.tree.grid(row=2,column=0, columnspan=5,ipadx=150,padx=10, pady=5)
        self.ActualizarTree()

        

    def getFullPeriodoZun(self,id_periodo): 
        
        try:            
            queryPeriodo = "SELECT id_peri, nombre, fecha_inicio, fecha_fin, orden  FROM ZUNpr.dbo.n_periodo where id_peri ="+str(id_periodo)
            self.cursor.execute(queryPeriodo) 
            data = self.cursor.fetchone()         
            return data

        except Exception:
            messagebox.showerror("Error","Problema de conexión con la base de datos")
               

    def cargarcombo(self): 
        self.anno_trim=self.tx_anno.get()  
        if self.tx_anno.get() and self.tx_trimestre_name:
            options=[]
            try:            
                if self.anno_trim:
                    if self.tx_trimestre_name.get() == '1er Trimestre':
                        self.cursor.execute("SELECT id_peri,nombre,fecha_inicio  FROM ZUNpr.dbo.n_periodo where fecha_inicio like '%"+self.anno_trim+"%' and orden in (1,2,3)")
                    elif self.tx_trimestre_name.get() == '2do Trimestre':
                        self.cursor.execute("SELECT id_peri,nombre,fecha_inicio  FROM ZUNpr.dbo.n_periodo where fecha_inicio like '%"+self.anno_trim+"%' and orden in (4,5,6)")
                    elif self.tx_trimestre_name.get() == '3er Trimestre':
                        self.cursor.execute("SELECT id_peri,nombre,fecha_inicio  FROM ZUNpr.dbo.n_periodo where fecha_inicio like '%"+self.anno_trim+"%' and orden in (7,8,9)")
                    elif self.tx_trimestre_name.get() == '4to Trimestre':
                        self.cursor.execute("SELECT id_peri,nombre,fecha_inicio  FROM ZUNpr.dbo.n_periodo where fecha_inicio like '%"+self.anno_trim+"%' and orden in (10,11,12)")
                    else:    
                        self.cursor.execute("SELECT id_peri,nombre,fecha_inicio  FROM ZUNpr.dbo.n_periodo where fecha_inicio like '%"+self.anno_trim+"%'")
                else:
                    self.cursor.execute('SELECT id_peri,nombre,fecha_inicio  FROM ZUNpr.dbo.n_periodo')
                slist = self.cursor.fetchall()
                for row in slist:
                    options.append(str(row['id_peri'])+"-"+str(row['nombre']).rstrip()+"-"+str(row['fecha_inicio'])[:4])  

                self.cb_periodo['value']=options           

            except Exception:
                messagebox.showerror("Error","Problema de conexión con la base de datos")
        else:
            messagebox.showwarning("Verificar informción","Debe seleccionar el trimestre y el año")

    def addperiodo(self):   
        countElemTree = len(self.tree.get_children())
        if countElemTree != 3:      
            periodo = str(self.cb_periodo.get()).split('-')                
            if periodo[0] in self.periodoregistrado:
                messagebox.showwarning('Registro repetido','El período ya aparece registrado')
            elif self.cb_periodo.get() == '':
                messagebox.showwarning('Campo en blanco','Debe seleccionar un período')
            else:    
                self.tree.insert('','end',values=(periodo[0],periodo[1],self.getFullPeriodoZun(periodo[0])['fecha_inicio'],self.getFullPeriodoZun(periodo[0])['fecha_fin'],self.getFullPeriodoZun(periodo[0])['orden']))
                self.periodoregistrado.append(periodo[0])
        else:
            messagebox.showerror('Error de inserción', 'Solo se pueden agregar 3 meses al trimestre')


    def ActualizarTree(self):  
        self.periodoregistrado=[]          
        self.tree.delete(*self.tree.get_children())    
        queryPeriodo = 'SELECT * FROM postgres.public.periodo'               
        self.cursorLoc.execute(queryPeriodo)
        listPeriodo=self.cursorLoc.fetchall()  
        countResult = len(listPeriodo)
        if countResult != 0:
            self.registrado = True
        for row in  listPeriodo:
            self.tree.insert('','end',values=(row[0],row[1],row[3],row[4],row[2]))
            self.periodoregistrado.append(row[0])
           
        
    def reiniciarPeriodo(self):
        queryDperiodo = 'DELETE FROM postgres.public.periodo'
        queryUtilDist = 'DELETE FROM postgres.public.utilidades_distribucion'
        queryDEmp = 'DELETE FROM postgres.public.empleado'
        self.cursorLoc.execute(queryDperiodo)
        self.cursorLoc.execute(queryUtilDist)
        self.cursorLoc.execute(queryDEmp)
        self.connLoc.commit()
        self.ActualizarTree()
        self.registrado = False
    
    def save(self):
        trimestre_n=self.tx_trimestre_name.get()
        anno_t=self.tx_anno.get()
        periodo_sel=self.cb_periodo.get()
        countElemTree = len(self.tree.get_children())
        if countElemTree == 3:            
            try:
                if self.registrado is False:
                    query="INSERT INTO postgres.public.utilidades_distribucion (name_distribucionu,monto_distribuir,anno) \
                        VALUES ('"+trimestre_n+"',0,"+anno_t+") RETURNING id "
                    self.cursorLoc.execute(query)                
                    id_insert_util=str(self.cursorLoc.fetchone()[0])
                    for item in self.tree.get_children(): 
                        query1="INSERT INTO postgres.public.periodo (id,mes,ordent,fechainicio,fechafin)\
                            VALUES ("+str(self.tree.item(item)['values'][0])+",'"+str(self.tree.item(item)['values'][1])+"',"+str(self.tree.item(item)['values'][4])+",'"+str(self.tree.item(item)['values'][2])+"','"+str(self.tree.item(item)['values'][3])+"') \
                                RETURNING id"               
                        self.cursorLoc.execute(query1)                    
                        id_insert_periodo=str(self.cursorLoc.fetchone()[0])
                        query2="INSERT INTO postgres.public.utilidades_periodo_incluye (upincluye_utilidadesd_id,upincluye_periodo_id,efectuado) VALUES ("+id_insert_util+","+id_insert_periodo+",FALSE)" 
                        
                        self.cursorLoc.execute(query2)
                        self.connLoc.commit()
                        self.registrado = True
                        messagebox.showinfo('Confirmación','El período se registró satisfactoriamente')
                else:
                    messagebox.showinfo("Verificando información",'Ya existe un trimestre registrado, reinicie el período si desea iniciar otro trimestre')

            except Exception as error:
                messagebox.showerror("Error",error)
        else:
            messagebox.showwarning('Datos incompletos', 'Debe completar 3 meses del trimestre')
    