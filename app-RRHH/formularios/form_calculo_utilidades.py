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
            #variablesde estado
            self.registro_salario = False
            self.registro_vacaciones = False
            # Definiendo controles de seleccion
            self.tx_empleado = ttk.Entry(panel_principal, font=('Times', 14), width=10)
            self.tx_empleado.grid(row=0,column=0,padx=5,pady=5,ipadx=40)

            #Boton para buscar empleados        
            self.btn_bempleados = tk.Button(panel_principal, text="Buscar", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.actualizartreeEUtil)
            self.btn_bempleados.place(x=250, y=2)
            
            
            #Label mostrasr total de registros
            self.tx_total = tk.Label(panel_principal, font=('Times', 18), bg=COLOR_CUERPO_PRINCIPAL, text='Total de registros: 0')
            self.tx_total.place(x=750, y=140)
            
            #Para buscar por departamento
            #Label departamento
            self.tx_departamento = tk.Label(panel_principal, font=('Times', 14), width=20, bg=COLOR_CUERPO_PRINCIPAL, text='Departamento:')
            self.tx_departamento.place(x=350, y=5)

            #Combo departamento
            self.cb_departamento= ttk.Combobox(panel_principal, width=30)
            self.cb_departamento.bind('<<ComboboxSelected>>', self.actualizartreeEUtil)
            #self.cb_periodo.current(0)
            self.cb_departamento.place(x=520, y=5)

            #Boton para agregar eva        
            self.btn_agEva = tk.Button(panel_principal, text="Registrar salario", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.regSal)
            self.btn_agEva.place(x=750, y=100)

            #Boton para agregar eva        
            self.btn_saveEva = tk.Button(panel_principal, text="Registrar vacaciones", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.regVac)
            self.btn_saveEva.place(x=750, y=180)

            #Boton aprobar evaluaciones        
            self.btn_signEva = tk.Button(panel_principal, text="Mostrar resumen", font=(
                'Times', 13), bg=COLOR_BARRA_SUPERIOR, bd=0, fg=COLOR_CUERPO_PRINCIPAL, command=self.showResumen)
            self.btn_signEva.place(x=750, y=230)                 
            

            #Treeview        
            columns = ('numero', 'nombreap', 'ci', 'salario','vacaciones','horast','coef','devengado')
            self.treeEUtil = ttk.Treeview(panel_principal, height=16, columns=columns, show='headings')
            self.style = ttk.Style(self.treeEUtil)
            self.style.configure('Treeview',rowheight=30)
            self.treeEUtil.column('numero',width=80)
            self.treeEUtil.column('nombreap',width=200)
            self.treeEUtil.column('ci',width=110)
            self.treeEUtil.column('salario',width=60)
            self.treeEUtil.column('vacaciones',width=60)
            self.treeEUtil.column('horast',width=60)
            self.treeEUtil.column('coef',width=60)
            self.treeEUtil.column('devengado',width=80)

            self.treeEUtil.heading(column='numero', text='No.')
            self.treeEUtil.heading(column='nombreap', text='Nombre y apellidos')
            self.treeEUtil.heading(column='ci', text='CI')
            self.treeEUtil.heading(column='salario', text='Mt. Sal')
            self.treeEUtil.heading(column='vacaciones', text='Mt. Vac')
            self.treeEUtil.heading(column='horast', text='Horas T.')
            self.treeEUtil.heading(column='coef', text='C. Eva')
            self.treeEUtil.heading(column='devengado', text='S. Dev')
            self.treeEUtil.grid(row=1,column=0, columnspan=5,ipadx=5,padx=5,pady=5)
            self.actualizartreeEUtil() 
            self.cargarDpto() 
            self.resumendetallemin()
            #print(self.getVacacionesMT('0091'))
        else:
            messagebox.showinfo('Notificación','Debe registrar un período de evaluación')


    #Definiendo tree view de periodo
    def regSal(self):                
        queryEmpL=''
        countReg = 0
        self.limpiarNominaLoc()
        queryEmpL='SELECT id FROM postgres.public.empleado'
        
        self.cursorLoc.execute(queryEmpL)
        slistEmp = self.cursorLoc.fetchall()            
        for row in slistEmp:
            #Cargar nomina del primer mes del periodo
            queryGetNom1 = "SELECT x.Id_nomina_sal,x.no_interno,x.deveng_salario,x.pago_2,dias_lab FROM ZUNpr.dbo.nomina_sal x where x.no_interno="+row[0]+" and x.id_periodo="+str(self.getPeriodo()[0][0])
            self.cursorZun.execute(queryGetNom1)
            sal_emp=self.cursorZun.fetchone()
            if sal_emp is not None:            
                queryInsertSalLoc = "INSERT INTO postgres.public.pago_salario (id, sal_devengado,	destajo,	horast,	psalario_empleado_id,	psalario_periodo_id) VALUES("+str(sal_emp['Id_nomina_sal'])+","+str(sal_emp['deveng_salario'])+","+str(sal_emp['pago_2'])+","+str(sal_emp['dias_lab'])+",'"+row[0]+"',"+str(self.getPeriodo()[0][0])+")"
                self.cursorLoc.execute(queryInsertSalLoc)
                self.connLoc.commit()
            #Cargar nomina del segundo mes del periodo
            queryGetNom1 = "SELECT x.Id_nomina_sal,x.no_interno,x.deveng_salario,x.pago_2,dias_lab FROM ZUNpr.dbo.nomina_sal x where x.no_interno="+row[0]+" and x.id_periodo="+str(self.getPeriodo()[1][0])
            self.cursorZun.execute(queryGetNom1)
            sal_emp=self.cursorZun.fetchone()
            if sal_emp is not None:
                queryInsertSalLoc = "INSERT INTO postgres.public.pago_salario (id, sal_devengado,	destajo,	horast,	psalario_empleado_id,	psalario_periodo_id) VALUES("+str(sal_emp['Id_nomina_sal'])+","+str(sal_emp['deveng_salario'])+","+str(sal_emp['pago_2'])+","+str(sal_emp['dias_lab'])+",'"+row[0]+"',"+str(self.getPeriodo()[1][0])+")"
                self.cursorLoc.execute(queryInsertSalLoc)
                self.connLoc.commit()
            #Cargar nomina del tercer mes del periodo
            queryGetNom1 = "SELECT x.Id_nomina_sal,x.no_interno,x.deveng_salario,x.pago_2,dias_lab FROM ZUNpr.dbo.nomina_sal x where x.no_interno="+row[0]+" and x.id_periodo="+str(self.getPeriodo()[2][0])
            self.cursorZun.execute(queryGetNom1)
            sal_emp=self.cursorZun.fetchone()
            if sal_emp is not None:
                queryInsertSalLoc = "INSERT INTO postgres.public.pago_salario (id, sal_devengado,	destajo,	horast,	psalario_empleado_id,	psalario_periodo_id) VALUES("+str(sal_emp['Id_nomina_sal'])+","+str(sal_emp['deveng_salario'])+","+str(sal_emp['pago_2'])+","+str(sal_emp['dias_lab'])+",'"+row[0]+"',"+str(self.getPeriodo()[2][0])+")"
                self.cursorLoc.execute(queryInsertSalLoc)
                self.connLoc.commit()
            countReg+=1
        self.tx_total['text'] = 'Total de registros: '+str(countReg)
        self.registro_salario = True
        messagebox.showinfo('Confirmación','La información de la nómina de salario se registró satisfactoriamente')

    def regVac(self):               
        queryEmpL='SELECT id FROM postgres.public.empleado'
        self.limpiarVacacionesLoc()
        self.cursorLoc.execute(queryEmpL)
        slistEmp = self.cursorLoc.fetchall()            
        for row in slistEmp:
            #Cargar vacaciones del mes-1 del periodo            
            id_inci=self.getVacaInciTrab(row[0],self.getPeriodo1()[0][0]) 
            if id_inci is not None:            
                queryVaca = "SELECT v.id_inci,v.tiempo_total,v.importe_total,v.dias_periodo,importe_periodo FROM ZUNpr.dbo.h_vacaciones v \
                WHERE v.id_inci ="+str(id_inci['id_inci'])  
                self.cursorZun.execute(queryVaca)
                dataVacaciones = self.cursorZun.fetchone()
                if dataVacaciones['dias_periodo'] == 0:
                    queryInsertVacaLoc = "INSERT INTO postgres.public.vacacionesp (id,dias,monto,vacacionesp_empleado_id,vacacionesp_periodo_id,tiempo_tota,importe_total) \
                    VALUES("+str(id_inci['id_inci'])+","+str(dataVacaciones['dias_periodo'])+","+str(dataVacaciones['importe_periodo'])+",'"+row[0]+"',"+str(self.getPeriodo1()[0][0])+","+str(dataVacaciones['tiempo_total'])+","+str(dataVacaciones['importe_total'])+")"
                    self.cursorLoc.execute(queryInsertVacaLoc)
                    self.connLoc.commit()
                if dataVacaciones['importe_total'] > dataVacaciones['importe_periodo'] and dataVacaciones['dias_periodo'] != 0:
                    importe_total = dataVacaciones['importe_total'] - dataVacaciones['importe_periodo']
                    tiempo_total = dataVacaciones['tiempo_total'] - dataVacaciones['dias_periodo']
                    queryInsertVacaLoc = "INSERT INTO postgres.public.vacacionesp (id,dias,monto,vacacionesp_empleado_id,vacacionesp_periodo_id,tiempo_tota,importe_total) \
                    VALUES("+str(id_inci['id_inci'])+",0,0,'"+row[0]+"',"+str(self.getPeriodo1()[0][0])+","+str(tiempo_total)+","+str(importe_total)+")"
                    self.cursorLoc.execute(queryInsertVacaLoc)
                    self.connLoc.commit()
                
            #Cargar vacaciones del mes 1 del periodo            
            id_inci=self.getVacaInciTrab(row[0],(self.getPeriodo1()[1][0]))                      
            if id_inci is not None:            
                queryVaca = "SELECT v.id_inci,v.tiempo_total,v.importe_total,v.dias_periodo,importe_periodo FROM ZUNpr.dbo.h_vacaciones v \
                WHERE v.id_inci ="+str(id_inci['id_inci'])
                self.cursorZun.execute(queryVaca)
                dataVacaciones = self.cursorZun.fetchone()
                if dataVacaciones['tiempo_total'] == 0:
                    queryInsertVacaLoc = "INSERT INTO postgres.public.vacacionesp (id,dias,monto,vacacionesp_empleado_id,vacacionesp_periodo_id,tiempo_tota,importe_total) \
                    VALUES("+str(id_inci['id_inci'])+","+str(dataVacaciones['dias_periodo'])+","+str(dataVacaciones['importe_periodo'])+",'"+row[0]+"',"+str(self.getPeriodo1()[1][0])+","+str(dataVacaciones['dias_periodo'])+","+str(dataVacaciones['importe_periodo'])+")"
                    self.cursorLoc.execute(queryInsertVacaLoc)
                    self.connLoc.commit()
                else:
                    queryInsertVacaLoc = "INSERT INTO postgres.public.vacacionesp (id,dias,monto,vacacionesp_empleado_id,vacacionesp_periodo_id,tiempo_tota,importe_total) \
                    VALUES("+str(id_inci['id_inci'])+","+str(dataVacaciones['dias_periodo'])+","+str(dataVacaciones['importe_periodo'])+",'"+row[0]+"',"+str(self.getPeriodo1()[1][0])+","+str(dataVacaciones['tiempo_total'])+","+str(dataVacaciones['importe_total'])+")"
                    self.cursorLoc.execute(queryInsertVacaLoc)
                    self.connLoc.commit()
            #Cargar vacaciones del mes 2 del periodo            
            id_inci=self.getVacaInciTrab(row[0],(self.getPeriodo1()[2][0]))            
            if id_inci is not None:   
                queryVaca = "SELECT v.id_inci,v.tiempo_total,v.importe_total,v.dias_periodo,importe_periodo FROM ZUNpr.dbo.h_vacaciones v \
                WHERE v.id_inci ="+str(id_inci['id_inci'])
                self.cursorZun.execute(queryVaca)
                dataVacaciones = self.cursorZun.fetchone()         
                if dataVacaciones['tiempo_total'] == 0:
                    queryInsertVacaLoc = "INSERT INTO postgres.public.vacacionesp (id,dias,monto,vacacionesp_empleado_id,vacacionesp_periodo_id,tiempo_tota,importe_total) \
                    VALUES("+str(id_inci['id_inci'])+","+str(dataVacaciones['dias_periodo'])+","+str(dataVacaciones['importe_periodo'])+",'"+row[0]+"',"+str(self.getPeriodo1()[2][0])+","+str(dataVacaciones['dias_periodo'])+","+str(dataVacaciones['importe_periodo'])+")"
                    self.cursorLoc.execute(queryInsertVacaLoc)
                    self.connLoc.commit()
                else:
                    queryInsertVacaLoc = "INSERT INTO postgres.public.vacacionesp (id,dias,monto,vacacionesp_empleado_id,vacacionesp_periodo_id,tiempo_tota,importe_total) \
                    VALUES("+str(id_inci['id_inci'])+","+str(dataVacaciones['dias_periodo'])+","+str(dataVacaciones['importe_periodo'])+",'"+row[0]+"',"+str(self.getPeriodo1()[2][0])+","+str(dataVacaciones['tiempo_total'])+","+str(dataVacaciones['importe_total'])+")"
                    self.cursorLoc.execute(queryInsertVacaLoc)
                    self.connLoc.commit()
            #Cargar vacaciones del mes 3 del periodo            
            id_inci=self.getVacaInciTrab(row[0],(self.getPeriodo1()[3][0]))            
            if id_inci is not None:                       
                queryVaca = "SELECT v.id_inci,v.tiempo_total,v.importe_total,v.dias_periodo,importe_periodo FROM ZUNpr.dbo.h_vacaciones v \
                WHERE v.id_inci ="+str(id_inci['id_inci'])
                self.cursorZun.execute(queryVaca)
                dataVacaciones = self.cursorZun.fetchone()
                if dataVacaciones['dias_periodo'] != 0:
                    queryInsertVacaLoc = "INSERT INTO postgres.public.vacacionesp (id,dias,monto,vacacionesp_empleado_id,vacacionesp_periodo_id,tiempo_tota,importe_total) \
                    VALUES("+str(id_inci['id_inci'])+","+str(dataVacaciones['dias_periodo'])+","+str(dataVacaciones['importe_periodo'])+",'"+row[0]+"',"+str(self.getPeriodo1()[3][0])+","+str(dataVacaciones['dias_periodo'])+","+str(dataVacaciones['importe_periodo'])+")"
                    self.cursorLoc.execute(queryInsertVacaLoc)
                    self.connLoc.commit()
        self.registro_vacaciones = True
        messagebox.showinfo('Confirmación','La información de las vacaciones se registró satisfactoriamente')

    def limpiarNominaLoc(self):
        queryDemp = 'DELETE FROM postgres.public.pago_salario'
        self.cursorLoc.execute(queryDemp)
        self.connLoc.commit()

    def limpiarVacacionesLoc(self):
        queryDemp = 'DELETE FROM postgres.public.vacacionesp'
        self.cursorLoc.execute(queryDemp)
        self.connLoc.commit()

    def actualizartreeEUtil(self):
        self.treeEUtil.delete(*self.treeEUtil.get_children())         
        queryEmpL=''
        if self.tx_empleado.get() != '' and self.cb_departamento.get() == '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,rc.mtsalario,rc.mtvacaciones,rc.horastt,rc.coeficienteeva_utilidades,rc.devengado FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.id = a.id INNER JOIN postgres.public.resumen_calculo AS rc ON x.id = rc.resumen_empleado_id where x.nombreap like '%"+self.tx_empleado.get().upper()+"%' ORDER BY a.id ASC"
        elif self.cb_departamento.get() != '' and self.tx_empleado.get() == '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,rc.mtsalario,rc.mtvacaciones,rc.horastt,rc.coeficienteeva_utilidades,rc.devengado FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id INNER JOIN postgres.public.resumen_calculo_utilidades AS rc ON x.id = rc.resumen_empleado_id where a.area = '"+self.cb_departamento.get()+"' ORDER BY a.id ASC"
        elif self.tx_empleado.get() != '' and self.cb_departamento.get() != '':
            queryEmpL="SELECT x.id,x.nombreap,x.ci,rc.mtsalario,rc.mtvacaciones,rc.horastt,rc.coeficienteeva_utilidades,rc.devengado FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id INNER JOIN postgres.public.resumen_calculo_utilidades AS rc ON x.id = rc.resumen_empleado_id where x.nombreap like '%"+self.tx_empleado.get().upper()+"%' and a.area = '"+self.cb_departamento.get()+"' ORDER BY a.id ASC"
        else:
            queryEmpL='SELECT x.id,x.nombreap,x.ci,rc.mtsalario,rc.mtvacaciones,rc.horastt,rc.coeficienteeva_utilidades,rc.devengado FROM postgres.public.empleado AS x INNER JOIN postgres.public.area AS a ON x.empleado_area_id = a.id INNER JOIN postgres.public.resumen_calculo_utilidades AS rc ON x.id = rc.resumen_empleado_id ORDER BY a.id ASC'
               
        self.cursorLoc.execute(queryEmpL)

        slistEmp = self.cursorLoc.fetchall()       
        for row in slistEmp:
            self.treeEUtil.insert('','end',values=("'"+row[0]+"'",row[1],row[2],row[3],row[4],row[5],row[5]))
        self.tx_total['text'] = 'Total de registros: '+str(len(self.treeEUtil.get_children()))
            

    def cargarDpto(self):
        options=[]         
        queryP='SELECT x.* FROM postgres.public.area x order by area asc'
        self.cursorLoc.execute(queryP)
        slistArea=self.cursorLoc.fetchall()
        for row in slistArea:
            options.append(row[1])
        
        self.cb_departamento['values']=options

    def getPeriodo(self):         
        queryP='SELECT p.* FROM postgres.public.utilidades_periodo_incluye x INNER JOIN postgres.public.periodo AS p ON x.upincluye_periodo_id = p.id order by p.id asc'
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()

    def getPeriodo1(self):         
        queryP='SELECT p.* FROM postgres.public.periodo AS p order by p.id asc'
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()

    def obtenerPerMes(self,mes):
        queryP="SELECT * FROM postgres.public.periodo where mes='"+str(mes)+"'"
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchall()
   
    def getVacaInciTrab(self,no_interno, periodo):
        queryNomina = "SELECT no_interno FROM ZUNpr.dbo.nomina_vac WHERE id_periodo = "+str(periodo)+" AND no_interno = "+no_interno+" AND deveng_vac != 0"
        self.cursorZun.execute(queryNomina)        
        idEmplwNom = self.cursorZun.fetchone()
        if idEmplwNom is not None:
            queryInci = "SELECT x.id_inci,x.id_padre FROM ZUNpr.dbo.h_incidencias x \
            WHERE x.no_interno = '"+no_interno+"' AND x.id_ppago ="+str(periodo)+" AND x.tipo = 4 ORDER BY x.id_inci DESC"
            self.cursorZun.execute(queryInci)
            result = self.cursorZun.fetchone()
            if result is not None:
                if result['id_padre']==0:
                    return result    

    def getDepartamento(self,idemp):         
        queryP="SELECT a.area  FROM postgres.public.empleado emp INNER JOIN postgres.public.area AS a ON emp.empleado_area_id  = a.id where emp.id = "+str(idemp)
        self.cursorLoc.execute(queryP)
        return self.cursorLoc.fetchone()[0]

    def showResumen(self): 
        queryP="SELECT emp.id,emp.nombreap,emp.ci,a.area  FROM postgres.public.empleado emp INNER JOIN postgres.public.area AS a ON emp.empleado_area_id  = a.id"
        self.cursorLoc.execute(queryP)
        empList = self.cursorLoc.fetchall()
        for emp in empList:
            mtvacaciones = 0
            mtsalario = 0
            horast = 0
            listSal = self.getpagoSalMT(emp[0])
            listVaca = self.getVacacionesMT(emp[0])
            for rowsal in listSal:
                mtsalario += rowsal[0]
                horast += rowsal[1]
            for rowvaca in listVaca:
                mtvacaciones += rowvaca[0]
            queryInsertRe = "INSERT INTO postgres.public.resumen_calculo_utilidades\
                (resumen_empleado_id,resumen_utilidadesd_id,mtvacaciones,mtsalario,horastt,coeficienteeva_utilidades,descrip_coeficiente,devengado)\
                    VALUES ('"+emp[0]+"',"+str(self.getUtiliDist()[0])+","+str(round(mtvacaciones,2))+","+str(round(mtsalario,2))+","+str(horast)+","+str(self.calcCoeficienteEva(emp[0]))+",'',"+str(round((mtsalario+mtvacaciones)))+")"
            self.cursorLoc.execute(queryInsertRe)
            self.connLoc.commit()
        self.actualizartreeEUtil()

    #Obtener informacion del Periodo de utilidades definido        
    def getUtiliDist(self):
        queryUD="SELECT x.* FROM postgres.public.utilidades_distribucion x"
        self.cursorLoc.execute(queryUD)
        return self.cursorLoc.fetchone()        

    def calcCoeficienteEva(self, emp):
        listPer = self.getPeriodo()
        sumeva = 0
        countDiv = len(listPer)
        for periodo in listPer:
            queryEva = "SELECT te.peso FROM postgres.public.tipo_evaluacion te \
            INNER JOIN postgres.public.evaluacion AS e ON te.id = e.evaluacion_tipoevaluacion_id WHERE e.evaluacion_perio_id="\
            +str(periodo[0])+" AND e.evaluacion_empleado_id = '"+emp+"'"
            self.cursorLoc.execute(queryEva)
            eva = self.cursorLoc.fetchone()
            if eva is not None:
                if eva[0] == 0:
                    countDiv -= 1
                else:
                    sumeva += eva[0]
            else:
                messagebox.showinfo('Notificación','No tiene registro de evaluación en el mes de '+str(periodo[1]))
        return sumeva/countDiv

    def resumendetallemin(self):
        path = "file/resumen_detalle_min.xlsx"
        row = 3 
        controw = 1
        self.limpiarExcel(row,path)   
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
             
        queryP="SELECT a.area,emp.id,emp.ci,emp.nombreap,emp.escalas,emp.thoraria  FROM postgres.public.empleado emp INNER JOIN postgres.public.area AS a ON emp.empleado_area_id  = a.id ORDER BY a.id"
        self.cursorLoc.execute(queryP)
        listEmp = self.cursorLoc.fetchall()
        for empleado in listEmp:
            sheet['B'+str(row)] =  controw
            sheet['C'+str(row)] =  empleado[1]
            sheet['D'+str(row)] =  empleado[2]
            sheet['E'+str(row)] =  empleado[3]
            sheet['F'+str(row)] =  empleado[4]
            sheet['G'+str(row)] =  empleado[5]
            if self.getDestajo(empleado[1])[1]:
                sheet['H'+str(row)] =  'Si'
            else:
                sheet['H'+str(row)] =  'No'
            vacaciones = self.getVacacionesMT(empleado[1])
            for v in vacaciones:
                idsperiodos =  []
                periodos = list(self.getPeriodo())
                idsperiodos.append(periodos[0][0])
                idsperiodos.append(periodos[1][0])
                idsperiodos.append(periodos[2][0])
                sheet['K'+str(row)] =  '0'
                sheet['L'+str(row)] =  '0'
                sheet['S'+str(row)] =  '0'
                sheet['T'+str(row)] =  '0'
                sheet['AA'+str(row)] =  '0'
                sheet['AB'+str(row)] =  '0'
                if v[2] in idsperiodos:
                    if idsperiodos.index(v[2]) == 0:
                        sheet['K'+str(row)] =  v[1]
                        sheet['L'+str(row)] =  v[0]
                    if idsperiodos.index(v[2]) == 1:
                        sheet['S'+str(row)] =  v[1]
                        sheet['T'+str(row)] =  v[0]
                    if idsperiodos.index(v[2]) == 2:
                        sheet['AA'+str(row)] =  v[1]
                        sheet['AB'+str(row)] =  v[0]
                else:
                    sheet['I'+str(row)] =  v[1]
                    sheet['J'+str(row)] =  v[0]
            salarios = self.getpagoSalMT(empleado[1])
            for s in salarios:
                sheet['M'+str(row)] =  s[2]
                sheet['N'+str(row)] =  s[0]
                sheet['U'+str(row)] =  s[2]
                sheet['V'+str(row)] =  s[0]
                sheet['AC'+str(row)] =  s[2]
                sheet['AD'+str(row)] =  s[0]

            row += 1
            controw += 1
        wb.save(path)



    def limpiarExcel(self,fila,url):         
        path = url
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        sheet.delete_rows(fila, sheet.max_row-1)        
        wb.save(path)




    def getpagoSalMT(self, empleado):
        mtsalario = []
        querygetSal="SELECT ps.sal_devengado,ps.destajo,ps.horast FROM postgres.public.utilidades_periodo_incluye up \
        INNER JOIN postgres.public.pago_salario AS ps ON up.upincluye_periodo_id = ps.psalario_periodo_id where ps.psalario_empleado_id  = '"+empleado+"' order by up.upincluye_periodo_id asc" 
        self.cursorLoc.execute(querygetSal)
        salarioList = self.cursorLoc.fetchall()
        for sal in salarioList:
            mtsalario.append(((sal[0] - sal[1]),sal[2],sal[1]))
        return mtsalario


    def getVacacionesMT(self, empleado):
        mtvacaciones = []
        querygetVaca="SELECT v.importe_total,v.tiempo_tota,p.id  FROM postgres.public.periodo p \
        INNER JOIN postgres.public.vacacionesp AS v ON p.id = v.vacacionesp_periodo_id where v.vacacionesp_empleado_id  = '"+empleado+"' order by p.id asc" 
        self.cursorLoc.execute(querygetVaca)
        vacacionesList = self.cursorLoc.fetchall()
        for vac in vacacionesList:
            if self.getDestajo(empleado)[1]:
                importe=float(vac[0])
                tarifaH=float(self.getDestajo(empleado)[0])
                calcVaca = importe*tarifaH*8
                mtvacaciones.append((calcVaca, vac[1],vac[2]))
            else:
                mtvacaciones.append((vac[0], vac[1],vac[2]))
        return mtvacaciones

        

        

    def getDestajo(self,empleado):
        querygetSal="SELECT e.thoraria,e.destajo FROM postgres.public.empleado AS e \
            where e.id  = '"+str(empleado)+"'"
        self.cursorLoc.execute(querygetSal)
        result = self.cursorLoc.fetchone()
        return result


    

    
