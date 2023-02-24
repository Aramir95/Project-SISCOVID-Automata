from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select 
from webdriver_manager.chrome import ChromeDriverManager
from unidecode import unidecode
import os
import logging
import time


class FailSiscovidBrowser(Exception):
    """Excepción para cuando no se puede iniciar el webdriver de Chrome

    Args:
        msg (str): Mensaje que debe mostrarse al encontrar una excepción de tipo FailSiscovidBrowser_
    """
    def __init__(self, msg):
        self.msg=msg
    
    def __str__(self):
        return self.msg

    

class SiscovidBrowser:
    def __init__(self, user: str = "", password: str = "",direccion_ipress:str="Av. Angamos Oeste 300, Miraflores, Lima, Lima") -> None:
        self.directory_path = self.find_path_chromedriver_directory()
        self.browser = self.inicializa_chrome()
        
        #Propiedades de usuario y contraseña y para el login
        self.user               = user
        self.password           = password
        self.direccion_empresa  = direccion_ipress
        self.xpath_buscar_paciente = {
            "seleccion_tipo_doc":"/html/body/div/div/div[1]/div/div/div/div[2]/div/div/div/div/div[2]/div[1]/div[1]/select",
            "input_buscar":'//*[@id="id_form_buscar-valor"]',
            "boton_buscar":'//*[@id="buscar"]'
        }
        
        self.visible_text_tipo_doc = {
            "DNI":"DNI",
            "CE":"Carnet de Extranjería",   #CARNET EXTRANJERÍA
            "PAS":"Pasaporte",  #PASAPORTE
            "CI":"Cedula de Identidad",   #CEDULA DE IDENTIDAD
            "CSR":"Carnet de solicitante de refugio",  #CARNET DE SOLICITANTE DE REFUGIO
            "ASI":"Sin Documento", #SIN DOCUMENTO
        }
        
        self.alertas = {"div_2":"/html/body/div[2]/div",
                       "h2_div_2":"/html/body/div[2]/div/h2",
                       "button_1":"/html/body/div[2]/div/div[10]/button[1]",
                       "button_2":"/html/body/div[2]/div/div[10]/button[2]",
                       "div_3":"/html/body/div[3]/div",
                       "if_error":"/html/body/div[3]/div/h2",
                       "boton_div3":"/html/body/div[3]/div/div[10]/button[1]",
                       "boton2_div3":"/html/body/div[3]/div/div[10]/button[2]",
                       "barra_buscar":"/html/body/div/div/div[1]/div/div/div/div[2]/div/div/div/div/div[1]/h3"
                       }
        
        self.xpath_formulario = {
            "fecha":    '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[1]/div[1]/input',
            "hora":     '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[1]/div[2]/input',
            "div":      '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div',
            "procedencia_establ_salud":'/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[4]/div/ul/li[2]/div/label',
            "tipo_lectura_visual":  '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[4]/div/div/ul/li[1]/div/label',
            "extranjero":           '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[5]/div/ul/li[4]/div/label',
            "contacto_confirmado":  '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[5]/div/ul/li[2]/div/label',
            "personal_salud":       '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[5]/div/ul/li[1]/div/label',
            "contacto_sospechoso":  '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[5]/div/ul/li[3]/div/label',
            "otro_priorizado":      "/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[5]/div/ul/li[5]/div/label",
            "input_otro_s":         "/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[2]/div/div/div/input",
            "fecha_inic_s":         "/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[3]/div/div/div/input",   
            "mayor_60":             "/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[7]/div/div/ul/li[1]/div/input" ,
            "ninguna_cond_r":       '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[7]/div/div/ul/li[14]/div/label',
            "clas_leve":            '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[6]/div/ul/li[2]/div/label',
            "clas_moderada":        '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[6]/div/ul/li[3]/div/label',
            "clas_severa":          '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[6]/div/ul/li[4]/div/label',
            "clas_asintomatico":    '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[6]/div/ul/li[1]/div/label',
            "hisopado_nasofaringeo":'/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[2]/div/div/ul/li[3]/div/label',
            "saliva":               '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[2]/div/div/ul/li[1]/div/label',
            "hisopado_nasal":       '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[2]/div/div/ul/li[2]/div/label',
            "hisopado_orofaringeo": '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[2]/div/div/ul/li[4]/div/label',
            "hisopado_nas_y_far":   '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[2]/div/div/ul/li[7]/div/label',
            "aspirado_traqueal":    '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[2]/div/div/ul/li[8]/div/label',
            "lavado_broncoalv":     '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[2]/div/div/ul/li[9]/div/label',
            "tejido_pulmonar":      '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[2]/div/div/ul/li[10]/div/label',
            "no_reactivo":          '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[5]/div/ul/li[1]/div/label',
            "reactivo":             '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/fieldset[3]/div[5]/div/ul/li[7]/div/label',
            "btn_enviar":           '/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[12]/button',
            "btn_f200":             '/html/body/div[3]/div/div[10]/button[2]'
            
            }
        
        self.ubicaciones_signos_checkbox = {
            "tos":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[1]/div/input",
            "garraspera":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[2]/div/input",
            "congestión":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[3]/div/input",
            "fiebre":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[4]/div/input",
            "dificultadr":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[5]/div/input",
            "diarrea":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[6]/div/input",
            "nausea":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[7]/div/input",
            "cefalea":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[8]/div/input",
            "mialgia":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[9]/div/input",
            "anosmia":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[10]/div/input",
            "ageusia":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[11]/div/input",
            "cansancio":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[12]/div/input",
            "fapetito":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[13]/div/input",
            "otro_s":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[14]/div/input",
            "ningun":"/html/body/div[1]/div/div/div[1]/div/div/div[2]/div/div[2]/form/div/div/div/div[2]/div[2]/div/div/div/div/div[3]/div/div[1]/ul/li[15]/div/input"
            }
        self.xpath_datos_pac={
            "nombre":       "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[2]/div[1]/div/div/input",
            "apellido_pat": "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[2]/div[2]/div/div/input",
            "apellido_mat": "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[2]/div[3]/div/div/input",
            "fecha_nac":    "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[2]/div[4]/input",
            "sexo":         "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[3]/div[1]/select",
            "etnia":        "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[3]/div[2]/select",
            "pais_nac":     "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[3]/div[3]/select",
            "tipo_seguro":  "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[3]/div[4]/div/select",
            "pais_proc":    "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[3]/div[6]/select",
            "cod_pais":     "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[4]/div[1]/select",
            "cod_pais_contacto":"/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[4]/div[3]/div[1]/select",
            "celular":      "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[4]/div[2]/input",
            "celular_contacto":"/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[4]/div[3]/div[2]/input",
            "correo":       "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[4]/div[4]/div/div/input",
            "informacion_domicilio":"/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[5]/div[1]/ul/li[1]/div/label",
            "lugar_donde_se_hospeda_actualmente":"/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[5]/div[1]/ul/li[2]/div/label",
            "input_geoloc": "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[7]/div[1]/div[2]/div[1]/textarea",
            "boton_geoloc": "/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[7]/div[1]/div[2]/div[2]/button",
            "map":"/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[7]/div[2]/div/div/div",
            "button_alert": "/html/body/div[3]/div/div[10]/button[1]",
            "div_alert":    "/html/body/div[3]/div",
            "select_distrito":"/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[6]/div[3]/select",
            "select_provincia":"/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[6]/div[2]/select",
            "select_departamento":"/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[6]/div[1]/select",
            "button_siguiente":"/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[11]/div/div/div/div[1]/button"
            }
        
        self.opciones_sexo={
            "M":"Masculino",
            "F":"Femenino"
            }

    def cerrar_navegador(self)->bool:
        """Función que cierra todas las ventanas del navegador creado en Siscovid Browser

        Returns:
            bool: True en caso se cierren todas las ventanas
        """
        try:
            self.browser.quit()
            return True
        except Exception as e:
            logging.error("Navegador ya se cerró.")
            return False
        
    def find_path_chromedriver_directory(self) -> str:
        """Busca el directorio de chromedriver_win32 en el path del proyecto"""
        directory_path=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                              "chromedriver_win32")
        return directory_path
    
    def inicializa_chrome(self) -> webdriver.Chrome :
        """Inicializa el webdriver de Chrome, si no se encuentra en el path, lo busca en el directorio 
        de chromedriver_win32 y lo instala.
        
        Returns:
            webdriver: Objeto webdriver de Chrome
            
        Raises:
            FailSiscovidBrowser: Si no se puede iniciar el webdriver de Chrome
        """
        # Genera la configuración para el webdriver de Chrome maximizado y sin extensiones para aumentar la velocidad
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument('--disable-extensions')
        
        try:
            browser = webdriver.Chrome(options=options)
            
        except Exception as e:
            
            try:
                # Excepción por falta de instalación de una version compatible de chromedriver.exe
                logging.error(f"{e.__class__}: {e}")
                browser = webdriver.Chrome(ChromeDriverManager(path=self.directory_path).install() , options=options)
            
            except Exception as e:
                # Excepción por cualquier otra razón que no sea falta de instalación de una version compatible de chromedriver.exe
                logging.error(f"SiscovidBrowser.inicializa_chrome: {e.__class__}: {e}")
                logging.error(f"{e.__class__}: {e}")
                raise FailSiscovidBrowser("No se pudo iniciar el webdriver de Chrome con SiscovidBrowser.inicializa_chrome()")
            
        del options
        
        return browser
    
    def Login_Siscovid(self) -> bool:
        """Inicia sesión en Siscovid"""
        
        login_xpath = {
            "boton_conglomerado":"/html/body/div/div/div/div/div/div/div/div[2]/div/form/div[2]/div/button",
            "username":'/html/body/div/div/div/div/div/div/div/div[2]/div/form/div[1]/input',
            "password":'/html/body/div/div/div/div/div/div/div/div[2]/div/form/div[2]/input',
            "boton_login":'//*[@id="botonEnviarLogin"]',
            "Busqueda de Paciente":'/html/body/div/div/div[1]/div/div/div/div[2]/div/div/div/div/div[1]/h3',
            "span_usuario_incorrecto":"/html/body/div/div/div/div/div/div/div/div[2]/div/form/div[3]/span",
            # Modificar en caso el usuario tenga más de un conglomerado asignado
            "id_place":'//*[@id="id_place"]',
            "modificar_select_place":'/html/body/div/div/div/div/div/div/div/div[2]/div/div[2]/div[1]/select/option[3]',
            "id_app":'//*[@id="id_app"]',
            "modificar_select_app":'/html/body/div/div/div/div/div/div/div/div[2]/div/div[2]/div[2]/select/option[3]',
            "boton_continuar":'//*[@id="btnContinuar"]'
            #############################################################           
                }
        
        def login_exitoso(self) -> bool:
            return self.find_element_by_xpath_and_compare_text(login_xpath.get("Busqueda de Paciente"),
                                                       "Busqueda de Paciente", 1)
        
     
        self.browser.get("https://siscovid.minsa.gob.pe/conglomerado/elegir_conglomerado/")
        self.click_clikeable_wait_xpath(login_xpath.get("boton_conglomerado"),3)
        self.input_text_wait_xpath(login_xpath.get("username"),self.user,3)
        self.input_text_instant_xpath(login_xpath.get("password"),self.password)
        self.click_instant_xpath(login_xpath.get("boton_login"))

            
        
        if not login_exitoso(self):
            
            try:
                self.click_clikeable_wait_xpath(login_xpath.get("id_place"),0.5)
                self.click_clikeable_wait_xpath(login_xpath.get("modificar_select_place"),2)
                self.click_clikeable_wait_xpath(login_xpath.get("id_app"),2)
                self.click_clikeable_wait_xpath(login_xpath.get("modificar_select_app"),2)
                self.click_instant_xpath(login_xpath.get("boton_continuar"))
                
                if not login_exitoso(self):
                    logging.error("Se ejecutaron todos los xpath pero no se inicializa el buscador de paciente")
                    logging.error(f"SiscovidBrowser.Login_Siscovid: {e.__class__}: {e}")
                    raise FailSiscovidBrowser("Login_Siscovid : No se pudo iniciar sesión en Siscovid")
                
                else:
                    return True    
            
            except Exception as e:
                
                if self.find_element_by_xpath_and_compare_text(login_xpath.get("span_usuario_incorrecto"),
                                                       "Error al ingresar con la credenciales proporcionadas", 1):
                    logging.error("USUARIO Y/O CONTRASEÑA INCORRECTOS")
                else:
                    logging.error("ES NECESARIO MODIFICAR EL CODIGO PARA QUE FUNCIONE CON SU USUARIO")
                logging.error(f"SiscovidBrowser.Login_Siscovid: {e.__class__}: {e}")
                raise FailSiscovidBrowser("Login_Siscovid : No se pudo iniciar sesión en Siscovid")
        
        else:
            return True
    


    def buscar_paciente(self,tipo_documento:str, num_documento:str):
        self.browser.get("https://siscovid.minsa.gob.pe/ficha/buscar/")
        self.wait_to_element_be_present(xpath=self.xpath_buscar_paciente.get("seleccion_tipo_doc"),time=10)
        self.select_element_by_visible_text(xpath=self.xpath_buscar_paciente.get("seleccion_tipo_doc"), visible_text=self.visible_text_tipo_doc.get(tipo_documento,"Sin Documento"))
        self.clear_box_xpath(xpath=self.xpath_buscar_paciente.get("input_buscar"))
        self.input_text_instant_xpath(xpath=self.xpath_buscar_paciente.get("input_buscar"), text=num_documento)
        self.click_instant_xpath(xpath=self.xpath_buscar_paciente.get("boton_buscar"))


    def fix_find_pacient(self):
        try:
            if self.if_element_xpath_exists(xpath=self.alertas.get("div_2"),time=1):
                if self.find_element_by_xpath_and_compare_text(
                    xpath=self.alertas.get("h2_div_2"),
                    expected_text="Alerta",
                    time=1
                    ):
                    self.click_instant_xpath(self.alertas.get("button_1"))
                    return False
                
                elif self.find_element_by_xpath_and_compare_text(
                    xpath=self.alertas.get("h2_div_2"),
                    expected_text="Paciente registrado en la Ficha00",
                    time=0.1
                    ):
                    self.click_instant_xpath(self.alertas.get("button_1"))
                    self.wait_to_element_be_not_present(self.alertas.get("div_2"))
                    return True
                
                elif self.find_element_by_xpath_and_compare_text(
                    xpath=self.alertas.get("h2_div_2"),
                    expected_text="Servicio de Reniec no responde",
                    time=0.1
                    ):
                    self.click_instant_xpath(xpath=self.alertas.get("button_2"))
                    self.wait_to_element_be_not_present(xpath=self.alertas.get("div_2"),time=1)
                    if self.if_element_xpath_exists(xpath=self.alertas.get("div_3"),time=1):
                        self.click_instant_xpath(xpath=self.alertas.get("boton2_div3"))
                        self.wait_to_element_be_not_present(xpath=self.alertas.get("div_3"))
                    return True

                else:
                    try:
                        self.click_instant_xpath(xpath=self.alertas.get("button_1"))
                    except Exception:
                        pass
                    return False

                
            elif self.if_element_xpath_exists(xpath=self.alertas.get("div_3")):
                #Alerta de Error
                if self.find_element_by_xpath_and_compare_text(xpath=self.alertas.get("if_error"),
                                                            expected_text="Error",time=0.1):
                    self.click_instant_xpath(self.alertas.get("boton_div3"))
                    self.wait_to_element_be_not_present(xpath=self.alertas.get("div_3"))
                    return False
                #Alerta no se encontraron datos de paciente
                else:
                    self.click_instant_xpath(self.alertas.get("boton_div3"))
                    self.wait_to_element_be_not_present(xpath=self.alertas.get("div_3"))
                    return True
                
            else:
                if self.find_element_by_xpath_and_compare_text(xpath=self.alertas.get("barra_buscar"),
                                                            expected_text="Busqueda de Paciente",
                                                            time=0.1):
                    return True
                else :
                    #Error de servidores
                    logging.error("ERROR DE SERVIDORES AL BUSCAR PACIENTE")
                    return False

        except Exception as e:
            logging.exception(e)
            logging.error(f"Error al consultar paciente en base de datos :{e.__class__}: {e}")
            return False
        
    def is_complete_data_paciente(self):
        if self.if_element_xpath_exists(xpath="/html/body/div[1]/div/div[1]/div/div/div/div[2]/div/div[2]/form/div[2]/div/div/div/div/div[4]/div[2]/input"):
            return False
        else:
            return True
    
    def llenar_formulario_prueba(self,Fecha,Hora_Ejecucion_de_la_prueba,Tiene_Sintomas,Fecha_de_inicio_de_Sintomas,
                                 Marque_los_Sintomas_presenta,Otros_especificar,Clasifica_clinica_severidad,
                                 Condicion_de_la_Persona,Tipo_Muestra,Resultado_de_la_prueba,OBSERVACION):
        try:
            #Ingresar Fecha y Hora de la prueba
            self.clear_and_input_text_xpath(xpath=self.xpath_formulario.get("fecha"),text=Fecha)
            self.clear_and_input_text_xpath(xpath=self.xpath_formulario.get("hora"),text=Hora_Ejecucion_de_la_prueba)
            self.click_instant_xpath(xpath=self.xpath_formulario.get("div"))
            #Modificar en caso se quiera cambiar la configuración o marcar una nueva opción
            #Procedencia de la muestra : Se ha configurado para marcar "De Establecimiento de Salud"
            self.click_instant_xpath(xpath=self.xpath_formulario.get("procedencia_establ_salud"))
            #Tipo de Lectura: Se ha configurado para marcar "Lectura Visual"
            self.click_instant_xpath(xpath=self.xpath_formulario.get("tipo_lectura_visual"))
            #Condición de la personas
            if Condicion_de_la_Persona == "Contacto con caso sospechoso".upper():
                self.click_instant_xpath(xpath=self.xpath_formulario.get("contacto_sospechoso"))
            elif Condicion_de_la_Persona =="Persona proveniente del extranjero (migraciones)".upper():
                self.click_instant_xpath(xpath=self.xpath_formulario.get("extranjero"))
            elif Condicion_de_la_Persona =="Contacto con caso confirmado".upper():
                self.click_instant_xpath(xpath=self.xpath_formulario.get("contacto_confirmado"))
            elif Condicion_de_la_Persona =="Personal de salud".upper():
                self.click_instant_xpath(xpath=self.xpath_formulario.get("personal_salud"))
            else :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("otro_priorizado"))
            
            #Deseleccionar checkbox de signos y sintomas
            for sign,xpath_s in self.ubicaciones_signos_checkbox.items():
                if self.is_check_box_selected(xpath_s):
                    xpath_label=xpath_s.replace("input","label")
                    self.click_instant_xpath(xpath_label)
                else :
                    pass
            
            if Tiene_Sintomas.replace("Í","I") == "SI":
                
                if Marque_los_Sintomas_presenta =="DOLOR DE GARGANTA":
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("garraspera").replace("input","label"))
                if Marque_los_Sintomas_presenta =="Dificultad respiratoria".upper():
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("dificultadr").replace("input","label"))
                elif Marque_los_Sintomas_presenta =="Fiebre".upper():
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("fiebre").replace("input","label"))
                elif Marque_los_Sintomas_presenta =="Congestión nasal".upper():
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("congestión").replace("input","label"))
                elif Marque_los_Sintomas_presenta =="Cefalea".upper():
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("cefalea").replace("input","label"))
                elif Marque_los_Sintomas_presenta =="Diarrea".upper():
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("diarrea").replace("input","label"))
                elif Marque_los_Sintomas_presenta =="Dolor articulaciones".upper():
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("mialgia").replace("input","label"))
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("otro_s").replace("input","label"))
                    self.input_text_wait_xpath(xpath=self.xpath_formulario.get("input_otro_s"),text="Dolor en las articulaciones")
                elif Marque_los_Sintomas_presenta =="Dolor Abdominal".upper():
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("mialgia").replace("input","label"))
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("otro_s").replace("input","label"))
                    self.input_text_wait_xpath(xpath=self.xpath_formulario.get("input_otro_s"),text="Dolor abdominal")
                elif Marque_los_Sintomas_presenta =="Dolor Muscular".upper():
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("mialgia").replace("input","label"))
                elif Marque_los_Sintomas_presenta =="Dolor Pecho".upper():
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("mialgia").replace("input","label"))
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("otro_s").replace("input","label"))
                    self.input_text_wait_xpath(xpath=self.xpath_formulario.get("input_otro_s"),text="Dolor en el pecho")
                elif Marque_los_Sintomas_presenta =="MALESTAR GENERAL":
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("mialgia").replace("input","label"))
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("congestión").replace("input","label"))
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("cansancio").replace("input","label"))
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("otro_s").replace("input","label"))
                    self.input_text_wait_xpath(xpath=self.xpath_formulario.get("input_otro_s"),text="Malestar General")
                else:
                    self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("otro_s").replace("input","label"))
                    self.input_text_wait_xpath(xpath=self.xpath_formulario.get("input_otro_s"),text=Marque_los_Sintomas_presenta)
                
                #Otros_signos_y_sintomas    
                if len(Otros_especificar)>2:
                    if self.is_check_box_selected(xpath=self.ubicaciones_signos_checkbox.get("otro_s")):
                        self.input_text_instant_xpath(xpath=self.xpath_formulario.get("input_otro_s"),text=f", {Otros_especificar}")
                    else:
                        self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("otro_s").replace("input","label"))
                        self.input_text_wait_xpath(xpath=self.xpath_formulario.get("input_otro_s"), text=Otros_especificar)
                
                #Fecha_de_inicio_de_síntomas
                self.clear_and_input_text_xpath(xpath=self.xpath_formulario.get("fecha_inic_s"),text=Fecha_de_inicio_de_Sintomas)
                self.click_instant_xpath(xpath=self.xpath_formulario.get("div"))
                
            else:
                self.click_instant_xpath(xpath=self.ubicaciones_signos_checkbox.get("ningun").replace("input","label"))
            
            #Condición de Riesgo
            #Ninguna condición de riesgo en caso de que no se marque por defecto ser mayor de 60 años
            if not self.is_check_box_selected(xpath=self.xpath_formulario.get("mayor_60")):
                self.click_instant_xpath(xpath=self.xpath_formulario.get("ninguna_cond_r"))
            
            #Clasificacion de severidad
            if Clasifica_clinica_severidad=="LEVE":
                self.click_instant_xpath(xpath=self.xpath_formulario.get("clas_leve"))
            elif Clasifica_clinica_severidad[:4]=="MODE" :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("clas_moderada"))
            elif Clasifica_clinica_severidad[:4]=="SEVE" :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("clas_severa"))
            elif Clasifica_clinica_severidad[:4]=="ASIN":
                self.click_instant_xpath(xpath=self.xpath_formulario.get("clas_asintomatico"))
            else:
                self.click_instant_xpath(xpath=self.xpath_formulario.get("clas_asintomatico"))
            
            #Tipo de Muestra
            if Tipo_Muestra == "Hisopado nasofaringeo".upper() :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("hisopado_nasofaringeo"))
            elif Tipo_Muestra == "Saliva".upper() :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("saliva"))
            elif Tipo_Muestra == "Hisopado nasal".upper() :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("hisopado_nasal"))
            elif Tipo_Muestra == "Hisopado orofaringeo".upper():
                self.click_instant_xpath(xpath=self.xpath_formulario.get("hisopado_orofaringeo"))
            elif Tipo_Muestra == "Hisopado nasal y faringeo".upper() :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("hisopado_nas_y_far"))
            elif Tipo_Muestra == "Aspirado Traqueal".upper() :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("aspirado_traqueal"))
            elif Tipo_Muestra == "Lavado Broncoalveolar".upper() :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("lavado_broncoalv"))
            elif Tipo_Muestra == "Tejido Pulmonar".upper() :
                self.click_instant_xpath(xpath=self.xpath_formulario.get("tejido_pulmonar"))
            else:
                self.click_instant_xpath(xpath=self.xpath_formulario.get("hisopado_nasofaringeo"))
            
            #Resultado de la prueba
            if Resultado_de_la_prueba == "NO REACTIVO":
                self.click_instant_xpath(xpath=self.xpath_formulario.get("no_reactivo"))
            elif Resultado_de_la_prueba == "REACTIVO":
                self.click_instant_xpath(xpath=self.xpath_formulario.get("reactivo"))
            else:
                raise FailSiscovidBrowser(f"Resultado de la prueba de  no es valido")
            
            #Ingresar Observación
            if len(OBSERVACION)>3:
                self.clear_and_input_text_xpath(xpath=self.xpath_formulario.get("observacion"),text=OBSERVACION)
            
            #ENVIAR
            self.click_instant_xpath(xpath=self.xpath_formulario.get("btn_enviar"))
            
            #OBVIAR F200
            if self.if_element_xpath_exists(xpath=self.xpath_formulario.get("btn_f200"),time=2):
                self.click_instant_xpath(xpath=self.xpath_formulario.get("btn_f200"))
            
        except Exception as e:
            logging.exception(e)
            logging.error(f" llenar_formulario_prueba(): {e.__class__}:{e}")
            raise FailSiscovidBrowser("Error al ejecutar función:  llenar_formulario_prueba()")   
                
    def wait_select_element_by_visible_text(self, xpath:str, visible_text:str, time:int=3):
        options_loaded = EC.presence_of_element_located((By.XPATH, f'{xpath}/option[1]'))
        WebDriverWait(self.browser, time).until(options_loaded)
        select=Select(self.browser.find_element(By.XPATH,xpath))
        select.select_by_visible_text(visible_text)
        selected_option = select.first_selected_option
        if selected_option.text == visible_text:
            return True
        else:
            logging.error("select_element_by_visible_text : No se seleccionó el elemento")
            raise FailSiscovidBrowser("select_element_by_visible_text no seleccionó el valor esperado")
        
    def select_element_by_visible_text(self, xpath:str, visible_text:str):
        select=Select(self.browser.find_element(By.XPATH,xpath))
        select.select_by_visible_text(visible_text)
        selected_option = select.first_selected_option
        if selected_option.text == visible_text:
            return True
        else:
            logging.error("select_element_by_visible_text : No se seleccionó el elemento")
            raise FailSiscovidBrowser("select_element_by_visible_text no seleccionó el valor esperado")
    
    def select_element_by_value(self, xpath: str, value: str):
        select = Select(self.browser.find_element(By.XPATH, xpath))
        select.select_by_value(value)
        selected_option = select.first_selected_option
        if selected_option.get_attribute("value") == value:
            return True
        else:
            logging.error("select_element_by_value: No se seleccionó el elemento")
            raise FailSiscovidBrowser("select_element_by_value no seleccionó el valor esperado")
    
    def is_check_box_selected(self,xpath:str):
        try:
            element_to_check=self.browser.find_element(By.XPATH,xpath).get_attribute('checked')
            if str(element_to_check) == "true":
                return True
            else:
                return False
        except Exception as e:
            logging.error(f"Error al verificar si el checkbox esta seleccionado :{e.__class__}: {e}")
            return False
    
    def if_element_xpath_exists(self,xpath:str, time:int = 0.5):
        try:
            self.wait_to_element_be_present(xpath=xpath, time=time)
            return True
        except Exception:
            return False

    def wait_to_element_be_present(self,xpath:str, time=5 ):
        WebDriverWait(self.browser,time)\
        .until(EC.presence_of_element_located((By.XPATH,xpath)))
        return True
    
    def wait_to_element_be_not_present(self, xpath:str, time=5):
        try:
            WebDriverWait(self.browser, time)\
                .until(EC.invisibility_of_element_located((By.XPATH, xpath)))
            return True
        except Exception:
            return False

    
    def click_clikeable_wait_xpath(self,xpath:str, time=5):
        WebDriverWait(self.browser,time)\
        .until(EC.element_to_be_clickable((By.XPATH,xpath)))\
        .click()
        return True
    
    def click_instant_xpath(self,xpath):
        self.browser.find_element(By.XPATH,xpath).click()
        return True
    
    def input_text_wait_xpath(self,xpath :str,text :str,time: int=5):
        WebDriverWait(self.browser,time)\
        .until(EC.element_to_be_clickable((By.XPATH,xpath)))
        self.browser.find_element(By.XPATH,xpath).send_keys(text)
        return True
    
    def input_text_instant_xpath(self,xpath,text):
        self.browser.find_element(By.XPATH,xpath).send_keys(text)
        return True
    
    def clear_box_xpath(self,xpath:str):
        self.browser.find_element(By.XPATH,xpath).clear()
    
    def clear_and_input_text_xpath(self,xpath:str,text:str):
        self.clear_box_xpath(xpath)
        self.input_text_instant_xpath(xpath,text)
        return True
    
    def find_element_by_xpath_and_compare_text(self, xpath:str, expected_text:str, time=10):
        try:
            # espera hasta que el elemento esté disponible en la página
            element = WebDriverWait(self.browser, time).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            # verifica si el contenido del elemento coincide con el texto esperado
            if element.text.strip() == expected_text:
                return True
            else:
                return False
        except:
            return False
    
    def if_input_is_read_only(self,xpath:str):
        try:
            element=self.browser.find_element(By.XPATH,xpath)
            if element.get_attribute("readonly"):
                return True
            else:
                return False
        except Exception:
            return False
    
    def input_text_for_not_read_only(self,text:str,xpath:str):
        [False if self.if_input_is_read_only(xpath) else self.clear_and_input_text_xpath(xpath=xpath, text=text)]
    


    def llenar_datos_paciente(self,Nombre:str, Apellido_Paterno:str, Apellido_Materno:str,
        Fecha_de_Nacimiento:str, Sexo:str, Etnia:str, Tipo_Seguro:str,
        Procedencia_pais:str, Codigo_Pais:str, Celular:str,Correo:str,Tipo_de_Residencia:str, Direccion:str,
        Departamento:str, Provincia:str, Distrito:str):
        
        try:
            direccion_completa=f"{Direccion}, {Departamento}, {Provincia}, {Distrito}"
            
            self.input_text_for_not_read_only(text=Nombre,xpath=self.xpath_datos_pac.get("nombre"))
            self.input_text_for_not_read_only(text=Apellido_Paterno,xpath=self.xpath_datos_pac.get("apellido_pat"))
            self.input_text_for_not_read_only(text=Apellido_Materno,xpath=self.xpath_datos_pac.get("apellido_mat"))
            self.input_text_for_not_read_only(text=Fecha_de_Nacimiento,xpath=self.xpath_datos_pac.get("fecha_nac"))
            
            
            if not self.if_input_is_read_only(xpath=self.xpath_datos_pac.get("sexo")):
                self.select_element_by_visible_text(xpath=self.xpath_datos_pac.get("sexo"),visible_text=self.opciones_sexo.get(Sexo))
            
            #Ingresar Etnia, por defecto en caso de no encontrar la etnia ingresada se selecciona Mestizo
            try:
                self.select_element_by_visible_text(xpath=self.xpath_datos_pac.get("etnia"),visible_text=Etnia.capitalize())
            except Exception:
                self.select_element_by_visible_text(xpath=self.xpath_datos_pac.get("etnia"),visible_text="Mestizo")
            
            #El tipo de seguro por defecto es Ninguno en la clase DataExcelSiscovid, esta parte de código junto a DATAEXCELSISCOVID debe modificarse para aceptar más tipos de seguro de requerirse
            try:
                self.select_element_by_visible_text(xpath=self.xpath_datos_pac.get("tipo_seguro"),visible_text=Tipo_Seguro.capitalize())
            except Exception:
                self.select_element_by_visible_text(xpath=self.xpath_datos_pac.get("tipo_seguro"),visible_text="Ninguno")
            
            #Establecer el país de nacionalidad y procedencia, por defecto es Perú
            try:
                self.select_element_by_visible_text(xpath=self.xpath_datos_pac.get("pais_nac"),visible_text=Procedencia_pais)
                self.select_element_by_visible_text(xpath=self.xpath_datos_pac.get("pais_proc"),visible_text=Procedencia_pais)
            except Exception:
                self.select_element_by_visible_text(xpath=self.xpath_datos_pac.get("pais_nac"),visible_text="PERU")
                self.select_element_by_visible_text(xpath=self.xpath_datos_pac.get("pais_proc"),visible_text="PERU")
            
            #Establecere codigo de pais, por defecto es 51, código país de número de contacto por defecto es 51
            try:
                self.select_element_by_value(xpath=self.xpath_datos_pac.get("cod_pais"),value=f"+{Codigo_Pais}")
            except Exception:
                self.select_element_by_value(xpath=self.xpath_datos_pac.get("cod_pais"),value="+51")
            self.select_element_by_value(xpath=self.xpath_datos_pac.get("cod_pais_contacto"),value="+51")
            
            #Ingresar número de celular, por defecto es 999909999 
            try:
                self.clear_and_input_text_xpath(xpath=self.xpath_datos_pac.get("celular"),text=Celular)
                self.clear_and_input_text_xpath(xpath=self.xpath_datos_pac.get("celular_contacto"),text=Celular) 
            except Exception:
                self.clear_and_input_text_xpath(xpath=self.xpath_datos_pac.get("celular"),text="999909999")
                self.clear_and_input_text_xpath(xpath=self.xpath_datos_pac.get("celular_contacto"),text="999909999")       
            
            #Ingresar correo, por defecto es no@gmail.com
            try:
                self.clear_and_input_text_xpath(xpath=self.xpath_datos_pac.get("correo"),text=Correo)
            except Exception:
                self.clear_and_input_text_xpath(xpath=self.xpath_datos_pac.get("correo"),text="no@gmail.com")
            
            #Seleccionar Tipo de residencia :
            try:
                if Tipo_de_Residencia=="Informacion de Domicilio".upper():
                    self.click_instant_xpath(self.xpath_datos_pac.get("informacion_domicilio"))
                elif Tipo_de_Residencia=="Lugar donde se hospeda actualmente".upper():
                    self.click_instant_xpath(self.xpath_datos_pac.get("lugar_donde_se_hospeda_actualmente"))
                else:
                    self.click_instant_xpath(self.xpath_datos_pac.get("informacion_domicilio"))
            except Exception:
                self.click_instant_xpath(self.xpath_datos_pac.get("lugar_donde_se_hospeda_actualmente"))
            
            #Ingresar dirección y geolocalización
            self.clear_and_input_text_xpath(xpath=self.xpath_datos_pac.get("input_geoloc"),text=direccion_completa)
            self.click_instant_xpath(self.xpath_datos_pac.get("boton_geoloc"))
            
            
            if self.if_element_xpath_exists(xpath=self.xpath_datos_pac.get("button_alert"),time=1):
                self.click_instant_xpath(self.xpath_datos_pac.get("button_alert"))
                self.try_put_location_ipress()
            
            else:
                self.verify_department_provincy_district(departamento=Departamento,provincia=Provincia,distrito=Distrito) 
            
            logging.warning("exito")
            
            #Enviar formulario
            self.click_instant_xpath(self.xpath_datos_pac.get("button_siguiente"))
            #Cerrar aviso de registro exitoso
            if self.if_element_xpath_exists(xpath=self.xpath_datos_pac.get("div_alert"),time=1.5):
                self.click_instant_xpath(xpath=self.xpath_datos_pac.get("button_alert"))
        
        except Exception as e:
            logging.exception(e)
            raise FailSiscovidBrowser("Error al registrar datos del paciente")

        

    
    def extract_select_options(self,xpath:str,time=3)-> dict:
        options_loaded = EC.presence_of_element_located((By.XPATH, f'{xpath}/option[1]'))
        WebDriverWait(self.browser, time).until(options_loaded)
        select = Select(self.browser.find_element(By.XPATH,xpath))
        options = select.options
        options_dict = {}
        for option in options:
            options_dict[option.get_attribute("value")] = unidecode(option.text.strip().upper())
            try:
                options_dict['selected_option'] = select.first_selected_option.get_attribute("value")
            except Exception:
                options_dict['selected_option'] = None
        return options_dict
    
    def verify_department_provincy_district(self,departamento:str,provincia:str,distrito:str):
        try:
            dict_departamento=self.extract_select_options(xpath=self.xpath_datos_pac.get("select_departamento"))
            if not dict_departamento.get(dict_departamento.get('selected_option'))==departamento:
                dict_departamento_invertido = {valor: llave for llave, valor in dict_departamento.items()}
                try:
                    self.select_element_by_value(xpath=self.xpath_datos_pac.get("select_departamento"),value=dict_departamento_invertido.get(departamento))
                except Exception:
                    respuesta=self.try_put_location_ipress()
                    return respuesta
                
            dict_provincia=self.extract_select_options(xpath=self.xpath_datos_pac.get("select_provincia"))
            if not dict_provincia.get(dict_provincia.get('selected_option'))==provincia:
                dict_provincia_invertido = {valor: llave for llave, valor in dict_provincia.items()}
                try:
                    self.select_element_by_value(xpath=self.xpath_datos_pac.get("select_provincia"),value=dict_provincia_invertido.get(provincia))
                except Exception:
                    respuesta=self.try_put_location_ipress()
                    return respuesta
            
            dict_distrito=self.extract_select_options(xpath=self.xpath_datos_pac.get("select_distrito"))
            if not dict_distrito.get(dict_distrito.get('selected_option'))==distrito:
                dict_distrito_invertido = {valor: llave for llave, valor in dict_distrito.items()}
                try:
                    self.select_element_by_value(xpath=self.xpath_datos_pac.get("select_distrito"),value=dict_distrito_invertido.get(distrito))
                except Exception:
                    respuesta=self.try_put_location_ipress()
                    return respuesta
            
            return True
            
        except Exception:
            try:
                self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_departamento"),visible_text="Lima")
                self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_provincia"),visible_text="Lima")
                self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_distrito"),visible_text="Lima")
                return True
            except Exception:
                return False
            
    
    def try_put_location_ipress(self):
        self.clear_and_input_text_xpath(xpath=self.xpath_datos_pac.get("input_geoloc"),text=self.direccion_empresa)
        self.click_instant_xpath(self.xpath_datos_pac.get("boton_geoloc"))
        
        distrito=self.direccion_empresa.split(",")[1].strip().upper()
        provincia=self.direccion_empresa.split(",")[2].strip().upper()
        departamento=self.direccion_empresa.split(",")[3].strip().upper()
        
        try:
            dict_departamento=self.extract_select_options(xpath=self.xpath_datos_pac.get("select_departamento"))
            if not dict_departamento.get(dict_departamento.get('selected_option'))==departamento:
                dict_departamento_invertido = {valor: llave for llave, valor in dict_departamento.items()}
                try:
                    self.select_element_by_value(xpath=self.xpath_datos_pac.get("select_departamento"),value=dict_departamento_invertido.get(departamento))
                except Exception:
                    self.select_element_by_value(xpath=self.xpath_datos_pac.get("select_departamento"),value=dict_departamento_invertido.get("LIMA"))
                    self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_provincia"),visible_text="Lima")
                    self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_distrito"),visible_text="Lima")
                    return True
                
            dict_provincia=self.extract_select_options(xpath=self.xpath_datos_pac.get("select_provincia"))
            if not dict_provincia.get(dict_provincia.get('selected_option'))==provincia:
                dict_provincia_invertido = {valor: llave for llave, valor in dict_provincia.items()}
                try:
                    self.select_element_by_value(xpath=self.xpath_datos_pac.get("select_provincia"),value=dict_provincia_invertido.get(provincia))
                except Exception:
                    self.select_element_by_value(xpath=self.xpath_datos_pac.get("select_departamento"),value=dict_departamento_invertido.get("LIMA"))
                    self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_provincia"),visible_text="Lima")
                    self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_distrito"),visible_text="Lima")
                    return True
            
            dict_distrito=self.extract_select_options(xpath=self.xpath_datos_pac.get("select_distrito"))
            if not dict_distrito.get(dict_distrito.get('selected_option'))==distrito:
                dict_distrito_invertido = {valor: llave for llave, valor in dict_distrito.items()}
                try:
                    self.select_element_by_value(xpath=self.xpath_datos_pac.get("select_distrito"),value=dict_distrito_invertido.get(distrito))
                    self.click_instant_xpath(self.xpath_datos_pac.get("lugar_donde_se_hospeda_actualmente"))
                except Exception:
                    self.select_element_by_value(xpath=self.xpath_datos_pac.get("select_departamento"),value=dict_departamento_invertido.get("LIMA"))
                    self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_provincia"),visible_text="Lima")
                    self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_distrito"),visible_text="Lima")
                    return True
            else:
                self.click_instant_xpath(self.xpath_datos_pac.get("lugar_donde_se_hospeda_actualmente"))
                return True
        
        except Exception:
            try:
                self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_departamento"),visible_text="Lima")
                self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_provincia"),visible_text="Lima")
                self.wait_select_element_by_visible_text(xpath=self.xpath_datos_pac.get("select_distrito"),visible_text="Lima")
                return True
            except Exception:
                return False
        
        
        
        
        
    