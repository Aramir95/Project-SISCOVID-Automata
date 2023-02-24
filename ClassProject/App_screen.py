from ClassProject.CloseWs import CloseWs
from ClassProject.DataExcel import DataExcel
from ClassProject.DataExcelSiscovid import DataExcelSiscovid
from ClassProject.SiscovidBrowser import SiscovidBrowser
from ClassProject.DataToRun import DataApp
from ClassProject.SiscovidRun import RunSiscovidFromData
import os
import time
import logging
import tkinter as tk
from tkinter import filedialog, Button, PhotoImage, Label
import tkinter.messagebox as messagebox
from PIL import Image, ImageTk
import shutil

     

class Screen_app:
    def __init__(self):
        self.ico_path=self.find_path_logo()
        self.report_xlsx_path=os.path.join(self.find_path_temp_directory(), "REPORTE.xlsx")
        self.revisar_xlsx_path=os.path.join(self.find_path_temp_directory(), "REVISAR.xlsx")
        self.report_docx_path=os.path.join(self.find_path_temp_directory(), "REPORTE.docx")
        self.app=self.window_init()
        self.window=None
        self.user_data=None
        self.frame_usuario= None
        self.frame__main=None
        self.exist_frame_usuario=False
        self.load_file_path=None
        self.status_siscovid=None
        self.width=None
        self.height=None
        
    def find_path_logo(self):
        ico_path=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                              "Images","logo.ico")
        return ico_path
    
    def find_path_temp_directory(self):
        temp_path=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                               "Temp")
        return temp_path
    
    def seleccionar_excel(self):
        CloseWs.close_excel()
        self.load_file_path=filedialog.askopenfilename(filetypes = [("Archivos de Excel", "*.xls;*.xlsx" )])
        objeto=DataExcelSiscovid(path=self.load_file_path)
        del objeto
        
        self.boton_seleccionar.config(text="Archivo Cargado", bg="#2ADB48", fg="#000000")
        return True
        
    def abrir_excel(self):
        CloseWs.close_excel()
        objeto=DataApp()
        objeto.open_excel()
        del objeto
        
        self.boton_abrir.config(text="Data Verificada", bg="#2ADB48", fg="#000000")
        
        return True
    
    def descargar_reporte(self):
        CloseWs.cerrar_excel_y_word()
        # Solicita al usuario que seleccione la carpeta de destino
        carpeta_destino = filedialog.askdirectory()
        #Copia los archivos temporales a la carpeta destino
        shutil.copy(self.report_xlsx_path, carpeta_destino)
        shutil.copy(self.revisar_xlsx_path, carpeta_destino)
        shutil.copy(self.report_docx_path, carpeta_destino)
        
        # Abre el archivo .docx
        os.startfile(os.path.join(carpeta_destino,"REPORTE.docx"))
        self.boton_guardar.config(text="Descargar último reporte",bg="#F2505D", fg="white")
        # Muestra un mensaje de descarga completada
        messagebox.showinfo(title="Descarga completada", message="Los archivos se han descargado en la carpeta seleccionada")
        
        return True
    
    def run_siscovid(self):
        user = self.user_data.get("usuario")
        password=self.user_data.get("contraseña")
        direccion=self.user_data.get("ubicación")
        
        self.status_siscovid=False
        self.boton_seleccionar.config(text="Seleccionar Excel",  bg="#F2505D", fg="white")
        self.boton_abrir.config(text="Verificar Data",  bg="#F2505D", fg="white")
        CloseWs.close_excel()
        try:
            App=RunSiscovidFromData(user=user, password=password, direccion_ipress=direccion)
            App.Ingresar_data_to_siscovid()
            App.cerrar_navegador()
            self.status_siscovid=True
            
            self.boton_guardar.config(text="Nuevo Reporte", bg="#DBC509", fg="#000000")
            
            return True
        except Exception as e:
            logging.error(e)
            try:
                App.cerrar_navegador()
            except Exception:
                pass
            messagebox.showerror(title="Error", message="No se pudo ejecutar el programa, verificar que los datos de usuario y contraseña sean correctos o que la data tenga el formato requerido.")
            self.status_siscovid=False
            return False
    
    def save_user_data(self):
        user_siscovid = {
            "usuario": self.usuario_entry.get(),
            "contraseña": self.contraseña_entry.get(),
            "ubicación": self.ubicación_entry.get(),
        }
        self.user_data=user_siscovid
        self.usuario_entry.delete(0, 'end')
        self.contraseña_entry.delete(0, 'end')
        self.ubicación_entry.delete(0, 'end')
        self.frame_usuario.destroy()
        self.exist_frame_usuario=False
    
    def toggle_password_visibility(self):
        if self.contraseña_entry.cget('show') == '':
            self.contraseña_entry.config(show='*')
        else:
            self.contraseña_entry.config(show='')
        
        if self.mostrar_button.cget('text') == 'Mostrar':
            self.mostrar_button.config(text='Ocultar')
            self.mostrar_button.config(bg="#F5C938")
        else:
            self.mostrar_button.config(text='Mostrar')
            self.mostrar_button.config(bg="#dc143c")
    
    def return_form_user(self):
        self.frame__main.destroy()
    
    
    def exit_app(self):
        self.window.quit()
    
    def window_init(self):
        self.window=tk.Tk()
        self.window.title("SISCOVID AUTOMATA PROJECT V-2.0 __ by: @AngelRamirezGomero")
        self.window.iconbitmap(self.ico_path)
        
        # Configuramos el tamaño de la ventana en términos de porcentajes
        self.width = round(self.window.winfo_screenwidth() * 0.85 ) # Ancho de la ventana en un 85% del ancho de la pantalla
        self.height = round(self.window.winfo_screenheight() * 0.80)  # Altura de la ventana en un 85% de la altura de la pantalla
        x = (self.window.winfo_screenwidth() // 2) - (self.width // 2)  # Centramos la ventana horizontalmente
        y = (self.window.winfo_screenheight() // 2) - (self.height // 2)  # Centramos la ventana verticalmente
        self.window.geometry('{}x{}+{}+{}'.format(self.width, self.height, x, y))
        self.window.configure(background="#373739")

        # Cargar la imagen y mostrarla en la parte superior de la ventana
        img = Image.open(self.ico_path.replace(".ico",".png")).resize((round(self.height*0.45),round(self.height*0.45)))
        img_tk = ImageTk.PhotoImage(img)
        label_img = tk.Label(self.window, image=img_tk, bg="#373739")
        label_img.place(relx=0.5, rely=0.235, anchor="center")
        #Texto debajo de la imagen
        text_title = tk.Label(self.window, text="   Automata Project V.2.0   ",bg="#373739", fg="#dc143c"
                              , height=1, font=("Helvetica", round(self.height*0.038),"bold") ).place(relx=0.5, rely=0.49, 
                                                                        anchor="center")

        self._frame=self.formulario_user()

        self.window.mainloop()
    
    def formulario_user(self):
        """Función que crea un frame con un formulario para ingresar los datos del usuario de SISCOVID
        y la dirección de la ipress en caso no se encuentre el domicilio para registrar al paciente.
        """
        self.exist_frame_usuario=True
        # Creamos el marco principal
        self.frame_usuario = tk.Frame(self.window, bg="#121212",borderwidth=2, relief="groove", padx=70,pady=0)
        self.frame_usuario.grid_columnconfigure(1, weight=1)
        self.frame_usuario.place(relx=0.5, rely=0.75, relwidth=0.7, relheight=0.45, anchor="center")
        

        #Creamos la etiquete de Formulario para ingresar Datos
        Title_label_frame=tk.Label(self.frame_usuario,text="INGRESE SUS CREDENCIALES SISCOVID", fg="#dc143c",bg="#121212",font=("Helvetica", round(self.height*0.03),"bold"))
        Title_label_frame.grid(row=0,column=0,columnspan=4, pady=round(self.height*0.03),sticky="nsew")
        tk.Label(self.frame_usuario, text="", bg="#121212", fg="white", font=("Helvetica", 1)).grid(row=1, column=0, pady=0, padx=10, sticky="w")
        # Creamos la etiqueta de usuario
        usuario_label = tk.Label(self.frame_usuario, text="Usuario:", bg="#121212", fg="white", font=("Helvetica", 12,"bold"))
        usuario_label.grid(row=2, column=0, pady=5, padx=10, sticky="w")
        
        # Creamos la entrada de usuario
        self.usuario_entry = tk.Entry(self.frame_usuario, width=30, justify="center", font=("Helvetica", 12, "bold"))
        self.usuario_entry.grid(row=2, column=1, pady=5, padx=10, sticky="nsew")
        
        
        
        # Creamos la etiqueta de contraseña
        contraseña_label = tk.Label(self.frame_usuario, text="Contraseña:", bg="#121212", fg="white", font=("Helvetica", 12,"bold"))
        contraseña_label.grid(row=3, column=0, pady=5, padx=10, sticky="w")
        
        # Creamos la entrada de contraseña
        self.contraseña_entry = tk.Entry(self.frame_usuario, width=30, font=("Helvetica", 12,"bold"), show="*", justify="center")
        self.contraseña_entry.grid(row=3, column=1, pady=5, padx=10, sticky="nsew")
        
        # Creamos el botón de mostrar contraseña
        self.mostrar_button = tk.Button(self.frame_usuario, text="Mostrar", command=self.toggle_password_visibility, bg="#dc143c", fg="white", font=("Helvetica", 10,"bold"), width=10)
        self.mostrar_button.grid(row=3, column=2, pady=5, padx=10, sticky="w")
        
        # Creamos la etiqueta de ubicación
        ubicación_label = tk.Label(self.frame_usuario, text="Dirección:", bg="#121212", fg="white", font=("Helvetica", 12,"bold"))
        ubicación_label.grid(row=4, column=0, pady=7, padx=10, sticky="w")
        # Creamos la entrada de ubicación
        self.ubicación_entry = tk.Entry(self.frame_usuario, width=50, font=("Helvetica", 12,"bold"), justify="center")
        self.ubicación_entry.grid(row=4, column=1, pady=7, padx=10, sticky="nsew")
        # Creamos la etiqueta de ejemplo
        ejemplo_label= tk.Label(self.frame_usuario,text="Dirección, Distrito, Provincia, Departamento (Ejemplo: Av. Arequipa 4857, Lince, Lima, Lima)", 
                                bg="#121212", fg="white", font=("Helvetica", round(self.height*0.012)), justify="left")
        ejemplo_label.grid(row=5,column=1,padx=10, pady=0, sticky="ew")
        
        # Creamos el botón de guardar
        guardar_button = tk.Button(self.frame_usuario, text="Guardar", command=self.save_user_data, width=10, height=1, bg="#dc143c", fg="white", font=("Helvetica", 11, "bold"))
        guardar_button.grid(row=6, column= 1, pady=10, padx=10, sticky="ew")

        
        self.frame_usuario.bind("<Destroy>", self.create_frame_utilities)

    def recreate_form_user(self, event):
        """Función que destruye el frame del formulario de usuario y crea uno nuevo"""
        self.formulario_user()
    

        
    def create_frame_utilities(self, event):
        
        self.frame__main = tk.Frame(self.window, bg="#F7F0E0",
                                        borderwidth=3, 
                                        relief="groove"
                                        , padx=50,pady=10
                                        )
        self.frame__main.grid_columnconfigure(1, weight=1)
        self.frame__main.place(relx=0.5, rely=0.75, relwidth=0.6, relheight=0.4, anchor="center")
        
        
        self.boton_seleccionar=tk.Button(self.frame__main, text="Seleccionar Excel", command=self.seleccionar_excel, width=10, height=1, bg="#F2505D", fg="white", font=("Helvetica", 13, "bold"))
        self.boton_seleccionar.grid(row=0, column= 1, pady=2, padx=10, sticky="nsew")
        #boton para abrir el exceL almacenado en la carpeta Data y cuya información será cargada a la plataforma SISCOVID
        self.boton_abrir=tk.Button(self.frame__main, text="Verificar Data", command=self.abrir_excel, width=10, height=1, bg="#F2505D", fg="white", font=("Helvetica", 13, "bold"))
        self.boton_abrir.grid(row=1, column= 1, pady=2, padx=10, sticky="nsew")
        
        self.boton_ejecutar=tk.Button(self.frame__main, text="Ejecutar SISCOVID", command=self.run_siscovid, width=10, height=1, bg="#F2505D", fg="white", font=("Helvetica", 13, "bold"))
        self.boton_ejecutar.grid(row=2, column= 1, pady=2, padx=10, sticky="nsew")
        
        self.boton_guardar=tk.Button(self.frame__main, text="Descargar último reporte", command=self.descargar_reporte, width=10, height=1, bg="#F2505D", fg="white", font=("Helvetica", 13, "bold"))
        self.boton_guardar.grid(row=3, column= 1, pady=2, padx=10, sticky="nsew")
        
        self.boton_regresar=tk.Button(self.frame__main, text="Reingresar USER", command=self.return_form_user, width=10, height=1, bg="#F2505D", fg="white", font=("Helvetica", 13, "bold"))
        self.boton_regresar.grid(row=4, column= 1, pady=2, padx=10, sticky="nsew")
        
        self.boton_salir=tk.Button(self.frame__main, text="Salir", command=self.exit_app, width=10, height=1, bg="#121212", fg="white", font=("Helvetica", 13, "bold"))
        self.boton_salir.grid(row=5, column= 1, pady=2, padx=10, sticky="nsew")
        
        self.frame__main.bind("<Destroy>", self.recreate_form_user)