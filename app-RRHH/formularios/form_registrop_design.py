import tkinter as tk
from tkinter import *
import pymssql
import psycopg2
from tkinter import ttk, messagebox
from config import COLOR_CUERPO_PRINCIPAL, COLOR_BARRA_SUPERIOR


class FormularioRegistroPDesign():

    def __init__(self, panel_principal):   
       
        #Definiendo variables
        self.var_periodo = StringVar()
        self.anno_trim = StringVar()
        self.opregistrado=[]
        self.periodoregistrado=[]
        # Definiendo controles 
        
        #comobo carga periodo del zun
        self.cbx_label = tk.Label(panel_principal, text="Período", bg=COLOR_CUERPO_PRINCIPAL)
        self.cbx_label.grid(row=1,column=0,padx=5,pady=10)
        
        self.cb_periodo = ttk.Combobox(panel_principal,textvariable=self.var_periodo, postcommand=self.cargarcombo)
        #self.cb_periodo.current(0)
        self.cb_periodo.grid(row=1,column=1,padx=5,pady=10)  
        

        #comobo carga de orden
        self.cbx_labelO = tk.Label(panel_principal, text="Orden", bg=COLOR_CUERPO_PRINCIPAL)
        self.cbx_labelO.grid(row=1,column=2,ipadx=20, padx=5,pady=10)
        
        self.cb_orden = ttk.Combobox(panel_principal,values=['1','2','3'], width=10)
        self.cb_orden.current(0)
        self.cb_orden.grid(row=1,column=3,padx=5,pady=10)

        #nombre del periodo
        self.tx_label = tk.Label(panel_principal, text="Nombre del período", bg=COLOR_CUERPO_PRINCIPAL)
        self.tx_label.grid(row=0,column=0,padx=5,pady=10)
        
        self.tx_trimestre_name = ttk.Entry(panel_principal, font=(
            'Times', 14), width=20)
        self.tx_trimestre_name.grid(row=0,column=1,padx=5,pady=10)

        #Año del periodo
        self.tx_labelA = tk.Label(panel_principal, text="Año", bg=COLOR_CUERPO_PRINCIPAL)
        self.tx_labelA.grid(row=0,column=2,padx=5,pady=10)
        
        self.annot = str(self.var_periodo)[:4]

        self.tx_anno = ttk.Entry(panel_principal, font=(
            'Times', 14), width=10)
        self.tx_anno.grid(row=0,column=3,padx=5,pady=10)

        #Boton para agregar periodo al treewiew
        self.btn_registro_periodo = tk.Button(panel_principal, text="Registrar", font=(
            'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, padx=5, command=self.addperiodo)
        self.btn_registro_periodo.grid(row=1,column=4,padx=5,pady=10)
        

        #Boton para crear Trimestre
        self.btn_save = tk.Button(panel_principal, text="Guargar", font=(
            'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, padx=15, command=self.save)
        self.btn_save.grid(row=3,columnspan=4,padx=10,pady=10)


        #Definiendo tree view de periodo

        self.tree = ttk.Treeview(panel_principal,
                                 show='headings')
        self.tree['columns'] = ('Id', 'Periodo', 'Orden')
        self.tree.column('Id',width=80)
        self.tree.column('Periodo',width=100)
        self.tree.column('Orden',width=100)

        self.tree.heading('Id', text='Id')
        self.tree.heading('Periodo', text='Período')
        self.tree.heading('Orden', text='Orden')
        self.tree.grid(row=2,column=0, columnspan=4,ipadx=150,padx=10, pady=5)

        

        

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
            

        except Exception:
            messagebox.showerror("Error","Problema de conexión con la base de datos")

    def addperiodo(self):        
        periodo = str(self.cb_periodo.get()).split('-')                
        if self.cb_orden.get() in self.opregistrado or periodo[0] in self.periodoregistrado:
            messagebox.showwarning('Registro repetido','El período o el orden ya aparece registrado')
        elif self.cb_periodo.get() == '':
            messagebox.showwarning('Campo en blanco','Debe seleccionar un período')
        else:    
            self.tree.insert('','end',values=(periodo[0],periodo[1],self.cb_orden.get()))
            self.opregistrado.append(self.cb_orden.get())
            self.periodoregistrado.append(periodo[0])
        

    
    def save(self):
        trimestre_n=self.tx_trimestre_name.get()
        anno_t=self.tx_anno.get()
        periodo_sel=self.cb_periodo.get()
        try:
            conn = psycopg2.connect(host="localhost", database="postgres", user="postgres", password="proyecto")
            cursor = conn.cursor()
            if trimestre_n =='' or anno_t=='' or periodo_sel=='' or self.periodoregistrado == []:
                messagebox.showinfo('Campos en blanco','Verifique, existen campos en blanco')
            else:
                query="INSERT INTO postgres.public.utilidades_distribucion (name_distribucionu,monto_distribuir,anno) \
                    VALUES ('"+trimestre_n+"',0,"+anno_t+") RETURNING id "
                cursor.execute(query)                
                id_insert_util=str(cursor.fetchone()[0])
                for item in self.tree.get_children(): 
                    query1="INSERT INTO postgres.public.periodo (id,mes,ordent)\
                          VALUES ("+str(self.tree.item(item)['values'][0])+",'"+str(self.tree.item(item)['values'][1])+"',"+str(self.tree.item(item)['values'][2])+")\
                            RETURNING id"               
                    cursor.execute(query1)                    
                    id_insert_periodo=str(cursor.fetchone()[0])
                    query2="INSERT INTO postgres.public.utilidades_periodo_incluye (upincluye_utilidadesd_id,upincluye_periodo_id,efectuado) VALUES ("+id_insert_util+","+id_insert_periodo+",FALSE)" 
                    
                    cursor.execute(query2)
                    conn.commit()

        except Exception as error:
            messagebox.showerror("Error",error)
    