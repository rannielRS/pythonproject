import tkinter as tk
from tkinter import font
from PIL import Image, ImageTk
from config import COLOR_BARRA_SUPERIOR, COLOR_MENU_LATERAL,COLOR_CUERPO_PRINCIPAL, COLOR_MENU_CURSOR_ENCIMA
import util.util_ventana as util_ventana
import util.util_imagenes as util_img

# Nuevo
from formularios.form_registrop_design import FormularioRegistroPDesign
from formularios.form_evaluacion_design import FormularioEvaluacionDesign
from formularios.form_cargaremp_design import FormularioCargarEDesign
from formularios.form_calculo_utilidades import FormularioCalcUtilidadesDesign
from formularios.form_op_design import FormularioOtrosPagosDesign

class FormularioMaestroDesign(tk.Tk):

    def __init__(self):
        super().__init__()
        self.logo = util_img.leer_imagen("./imagenes/inicio2.png", (750, 500))
        self.perfil = util_img.leer_imagen("./imagenes/rex.png", (250, 150))
        self.img_sitio_construccion = util_img.leer_imagen("./imagenes/sitio_construccion.png", (200, 200))
        self.config_window()
        self.paneles()
        self.controles_barra_superior()        
        self.controles_menu_lateral()
        self.controles_cuerpo()
    
    def config_window(self):
        # Configuración inicial de la ventana
        self.title('Sistema para el pago de las utilidades')
        img = tk.PhotoImage(file="./imagenes/logo1.png")  # Replace "image.png" with any image file.
        self.iconphoto(False, img)
        #self.iconbitmap("./imagenes/logo.ico")
        w, h = 1260, 620        
        util_ventana.centrar_ventana(self, w, h)        

    def paneles(self):        
         # Crear paneles: barra superior, menú lateral y cuerpo principal
        self.barra_superior = tk.Frame(
            self, bg=COLOR_BARRA_SUPERIOR, height=50)
        self.barra_superior.pack(side=tk.TOP, fill='both')  
        
        self.menu_lateral = tk.Frame(self, bg=COLOR_MENU_LATERAL, width=150)
        self.menu_lateral.pack(side=tk.LEFT, fill='both', expand=False)         
        
        self.cuerpo_principal = tk.Frame(
            self, bg=COLOR_CUERPO_PRINCIPAL)
        self.cuerpo_principal.pack(side=tk.LEFT, fill='both', expand=True) 
    
    def controles_barra_superior(self):
        # Configuración de la barra superior
        font_awesome = font.Font(family='FontAwesome', size=12)

        # Etiqueta de título
        self.labelTitulo = tk.Label(self.barra_superior, text="RRHH-REX")
        self.labelTitulo.config(fg="#fff", font=(
            "Roboto", 15), bg=COLOR_BARRA_SUPERIOR, pady=10, width=16)
        self.labelTitulo.pack(side=tk.LEFT)

        # Botón del menú lateral
        imagen_pil_btbl = Image.open("./imagenes/menud.png")
        imagen_resize_btbl = imagen_pil_btbl.resize((21,21))
        imagen_btbl_tk = ImageTk.PhotoImage(imagen_resize_btbl)
        #Boton barra lateral       
        self.buttonMenuLateral = tk.Button(self.barra_superior, text="\uf0c9", font=font_awesome, image=imagen_btbl_tk,
                                           command=self.toggle_panel, bd=0, bg=COLOR_BARRA_SUPERIOR)
        self.buttonMenuLateral.image = imagen_btbl_tk
        self.buttonMenuLateral.pack(side=tk.LEFT)

        # Etiqueta de informacion
        self.labelTitulo = tk.Label(
            self.barra_superior, text="informaticos@rex.cu")
        self.labelTitulo.config(fg="#fff", font=(
            "Roboto", 10), bg=COLOR_BARRA_SUPERIOR, padx=10, width=20)
        self.labelTitulo.pack(side=tk.RIGHT)
    
    def controles_menu_lateral(self):
        # Configuración del menú lateral
        ancho_menu = 20
        alto_menu = 2
        font_awesome = font.Font(family='FontAwesome', size=15)
         
         # Etiqueta de perfil
        self.labelPerfil = tk.Label(
            self.menu_lateral, image=self.perfil, bg=COLOR_CUERPO_PRINCIPAL)
        self.labelPerfil.pack(side=tk.TOP, pady=10)

        # Botones del menú lateral
        
        self.buttonRP = tk.Button(self.menu_lateral)        
        self.buttonIE = tk.Button(self.menu_lateral)        
        self.buttonET = tk.Button(self.menu_lateral)
        self.buttonDU = tk.Button(self.menu_lateral)        
        self.buttonOP = tk.Button(self.menu_lateral)
       

        buttons_info = [
            ("Registrar período", "\uf109", self.buttonRP,self.abrir_registrar_p),
            ("Importar empleados", "\uf007", self.buttonIE,self.abrir_cargarEmp),
            ("Evaluar trabajador", "\uf03e", self.buttonET,self.abrir_evaluacion),
            ("Distribuir utilidades", "\uf129", self.buttonDU,self.abrir_calc_util),
            ("Otros pagos", "\uf013", self.buttonOP,self.abrir_otros_pagos)
        ]

        for text, icon, button,comando in buttons_info:
            self.configurar_boton_menu(button, text, icon, font_awesome, ancho_menu, alto_menu,comando)                    
    
    def controles_cuerpo(self):
        # Imagen en el cuerpo principal
        label = tk.Label(self.cuerpo_principal, image=self.logo,
                         bg=COLOR_CUERPO_PRINCIPAL)
        label.place(x=0, y=0, relwidth=1, relheight=1)        
  
    def configurar_boton_menu(self, button, text, icon, font_awesome, ancho_menu, alto_menu, comando):
        button.config(text=f"{text}", anchor="w", font=font_awesome,
                      bd=0, bg=COLOR_MENU_LATERAL, fg="white", width=ancho_menu, height=alto_menu,
                      command = comando)
        button.pack(side=tk.TOP)
        self.bind_hover_events(button)

    def bind_hover_events(self, button):
        # Asociar eventos Enter y Leave con la función dinámica
        button.bind("<Enter>", lambda event: self.on_enter(event, button))
        button.bind("<Leave>", lambda event: self.on_leave(event, button))

    def on_enter(self, event, button):
        # Cambiar estilo al pasar el ratón por encima
        button.config(bg=COLOR_MENU_CURSOR_ENCIMA, fg='white')

    def on_leave(self, event, button):
        # Restaurar estilo al salir el ratón
        button.config(bg=COLOR_MENU_LATERAL, fg='white')

    def toggle_panel(self):
        # Alternar visibilidad del menú lateral
        if self.menu_lateral.winfo_ismapped():
            self.menu_lateral.pack_forget()
        else:
            self.menu_lateral.pack(side=tk.LEFT, fill='y')
    # Nuevo
    def abrir_registrar_p(self):   
        self.limpiar_panel(self.cuerpo_principal)     
        FormularioRegistroPDesign(self.cuerpo_principal)

    def abrir_cargarEmp(self):
        self.limpiar_panel(self.cuerpo_principal)     
        FormularioCargarEDesign(self.cuerpo_principal)
        
    def abrir_evaluacion(self):   
        self.limpiar_panel(self.cuerpo_principal)     
        FormularioEvaluacionDesign(self.cuerpo_principal) 

    def abrir_calc_util(self):  
        self.limpiar_panel(self.cuerpo_principal)          
        FormularioCalcUtilidadesDesign(self.cuerpo_principal)   

    def abrir_otros_pagos(self):  
        self.limpiar_panel(self.cuerpo_principal)          
        FormularioOtrosPagosDesign(self.cuerpo_principal) 

    def limpiar_panel(self,panel):
    # Función para limpiar el contenido del panel
        for widget in panel.winfo_children():
            widget.destroy()

    
