from ClassProject.DataExcel import DataExcel
from ClassProject.CloseWs import CloseWs
import pandas as pd
import numpy as np
import os
import logging
import re
from datetime import timedelta, datetime

class FailDataExcelSiscovid(Exception):
    """Genera una excepción cuando ocurre un error al cargar los datos del archivo de Excel.

    Args:
        msg (str): Mensaje que se desea mostrar cuando se genera la excepción.

    Attributes:
        msg (str): Mensaje que se desea mostrar cuando se genera la excepción.
    """
    def __init__(self, msg):
        self.msg=msg
    
    def __str__(self):
        return self.msg

class DataExcelSiscovid(DataExcel):
    """
    La clase `DataExcelSiscovid` se encarga de gestionar la carga de un archivo de Excel en formato .xls o .xlsx.
    Los archivos de excel trabajados serán transformados todos a formato texto para evitar errores de tipo de dato.
    Los dataframes generados por esta clase serán de tipo `pandas.DataFrame` pero con dtype `object` en todas las columnas.
    Esta clase esta diseñada para cargar los datos del archivo de Excel que almacenan la información necesaria para el registro
    de pruebas antigénicas contra el SARS-COV-2 en la plataforma SISCOVID de Perú.
    
    
    Atributos:
        path (str): Ruta del archivo Excel a cargar.
        header (str): Nombre de la cabecera, para este tipo de archivos será "Fecha".
        _frame (DataFrame): DataFrame resultante de cargar los datos del archivo, se sobreescribió atributo de la calse padre.
        save_path (str): Ruta del archivo de Excel en donde se guardará el DataFrame generado al inicializar la clase.

    Métodos heredados:
        path_verify() -> Verifica la ruta del archivo y la transforma a formato .xlsx en caso de ser necesario.
        load_excel() -> Carga los datos del archivo a un DataFrame.
        transform_xlsx_values_to_text(file_path:str) -> Transforma los valores de todas las celdas de un archivo .xlsx a formato de texto.
        transform_xls_to_xlsx(file_path:str) -> Transforma un archivo .xls a .xlsx y transforma los valores de todas las celdas a formato de texto.
    
    Métodos propios:
        data_to_run_path() -> Genera mediante la ubicación del archivo de ejecución, la ruta en donde se guardará el DataFrame generado al inicializar la clase.
        correction_data_Frame() -> Corrige algunos datos del DataFrame para que se ajusten a las necesidades de la plataforma SISCOVID.
    """
    #Genera el docstring
    def __init__(self, path:str, header:str = "Fecha", dtype_str: bool = True):
        """
         Inicializa la clase `DataExcelSiscovid` y carga los datos del archivo de Excel en un DataFrame.
         Para esto, inicializa la clase padre `DataExcel` y sobreescribe el atributo `frame` con el DataFrame generado por el método `correction_data_Frame()`.

        Args:
            path (str): Ruta del archivo Excel a cargar.
            header (str, opcional): Nombre de la cabecera a utilizar, en este clase será "Fecha". Por defecto, "Fecha".
            dtype_str(bool, opcional): Especifica si se quiere que todo el dataframe se lea como celdas con formato string.Por defecto True.
        """
        super().__init__(path, header, dtype_str)
        #Generación de la ruta en donde se guardará el DataFrame generado al inicializar la clase
        self.save_path = self.data_to_run_path()
        #Sobreescritura del atributo frame de la clase padre
        self._frame = self.correction_data_Frame()
    

    @property
    def frame(self) -> pd.DataFrame:
        """Propiedad que retorna el DataFrame generado por el método `correction_data_Frame()`.

        Returns:
            DataFrame: DataFrame generado por el método `correction_data_Frame()`."""   
        return self._frame
    
    
    def data_to_run_path(self):
        """Genera mediante la ubicación del archivo de ejecución, la ruta en donde se guardará el DataFrame generado al inicializar la clase.

        Returns:
            str: Ruta en donde se guardará el DataFrame generado al inicializar la clase.
        """
        path_run=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                              "Data","Data_to_run.xlsx")
        return path_run
   
   
    def correction_data_Frame(self):
        """Corrige algunos datos del DataFrame para que se ajusten a las necesidades de la plataforma SISCOVID.}
        Guarda el dataframe generado en el archivo de Excel ubicado en la ruta `self.save_path`.

        Returns:
            DataFrame: DataFrame generado por el método `correction_data_Frame()`.

        Raises:
            FailDataExcelSiscovid: Excepción que se genera cuando ocurre un error al cargar o corregir los datos del archivo de Excel.
        """       
            
        CloseWs.close_excel()
        
        try:
            dataframe= self.load_excel()

            dataframe=dataframe.fillna("")
            dataframe = dataframe.applymap(lambda x : x.strip() )
            dataframe = dataframe.apply(lambda x: x.str.upper() if x.dtype == "object" else x)
            dataframe=dataframe.reset_index(drop=True)
            dataframe.index = np.arange(1,len(dataframe)+1)
            
            columns = ['Fecha', 'Hora_Ejecucion_de_la_prueba', 'Tipo_Documento',
        'Nro_Documento', 'Nombre', 'Apellido_Paterno', 'Apellido_Materno',
        'Fecha_de_Nacimiento', 'Sexo', 'Etnia', 'Tipo_Seguro',
        'Procedencia_pais', 'Codigo_Pais', 'Celular','Correo', 'Tipo_de_Residencia', 'Direccion',
        'Departamento', 'Provincia', 'Distrito','Tiene_Sintomas', 'Fecha_de_inicio_de_Sintomas',
        'Marque_los_Sintomas_presenta', 'Otros_especificar',
        'Condicion_de_la_Persona', 
        'Procedencia_solicitud_diag', 'Resultado_de_la_prueba',
        'Clasifica_clinica_severidad', 'El_paciente_condicion_riesgo',
        'Tipo_Muestra', 'Tipo_Lectura','ESTADO_SISCOVID'
        ]

            for col in columns:
                if col not in dataframe.columns :
                    dataframe[col]=""
            
            #CREO COLUMNA OBSERVACIÓN PARA INICIAR DESDE "" con las observaciones encontradas
            dataframe=dataframe[columns]
            dataframe["OBSERVACIÓN"]=""

            
            def es_email_valido(email):
                pattern = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}$"
                if re.search(pattern, email):
                    email=email.replace("ggmail","gmail")
                    email=email.replace("mailc.om","mail.com")
                    return email
                else:
                    return "no_indica@gmail.com"

            def es_celular_valido(celular):
                pattern = r"^[0-9]{9}$"
                if re.search(pattern, celular):
                    if celular == "999999999": #No pueden ser todos los digitos iguales
                        return "909909909"
                    return celular
                else:
                    return "999000999"
            
        


            dataframe["Correo"]=dataframe["Correo"].str.lower()
            dataframe["Correo"]=dataframe["Correo"].apply(es_email_valido)
            dataframe["Celular"]=dataframe["Celular"].apply(es_celular_valido)
            
            #La aplicación fue diseñada con una base de datos de empresas privadas en donde el tipo de seguro
            #registrado siempre era ninguno, por lo que se asume que si el tipo de seguro es diferente a ninguno
            #entonces el tipo de seguro es ninguno. 
            #Modificar si se requiere para otras bases de datos o su uso en entornos mixtos.
            dataframe["Tipo_Seguro"]=np.where(
                dataframe["Tipo_Seguro"]!="NINGUNO",
                "NINGUNO",dataframe["Tipo_Seguro"])
            
            #Apellido Materno diferente a "-" y diferente a "". Problema con Extranjeros
            dataframe["Apellido_Materno"]=dataframe["Apellido_Materno"].str.replace("-",".")
            dataframe["Apellido_Materno"]=np.where(
                (dataframe["Apellido_Materno"] == "") | (dataframe["Apellido_Materno"]== "-")
                ,".",dataframe["Apellido_Materno"])
            
            #Corregir Dirección:
            dataframe["Direccion"]=np.where(
                (dataframe["Direccion"].str.len() < 5) | (dataframe["Direccion"]=="[NO ESPECIFICA]")
                , "LIMA", dataframe["Direccion"])
            
            #ESTADO SISCOVID ES NO INGRESADO PARA FILAS QUE NO TENGAN EL DATO
            dataframe["ESTADO_SISCOVID"] = np.where(
                dataframe["ESTADO_SISCOVID"] == "", "NO INGRESADO", dataframe["ESTADO_SISCOVID"])
            
            #CASO FECHA DE INICIO DE SINTOMAS, EN CASO TENGA SINTOMAS = SI Y LA FECHA DE INICIO DE SINTOMAS SEA "" SE LE PONE LA FECHA DE LA PRUEBA
            dataframe["Fecha_de_inicio_de_Sintomas"] = np.where(
                ((dataframe["Tiene_Sintomas"] == "SI") & (dataframe["Fecha_de_inicio_de_Sintomas"] =="" ))
                , dataframe["Fecha"], dataframe["Fecha_de_inicio_de_Sintomas"])
            
            # CASO INADMISIBEL: En caso DNI tenga más de 8 dígitos, o Carnet de extranjería tenga un número de dígitos diferente de 9 dígitos.
            dataframe["Nro_Documento"]=dataframe.apply(lambda x :  x.Nro_Documento.zfill(8) if x.Tipo_Documento=="DNI" else x.Nro_Documento , axis=1)
            
            dataframe["ESTADO_SISCOVID"] = np.where(
                ( (dataframe["Tipo_Documento"] == "DNI") & (dataframe["Nro_Documento"].str.len() != 8) ) |
                ( (dataframe["Tipo_Documento"] == "CE") & (dataframe["Nro_Documento"].str.len() != 9) )
                                                        , "INADMISIBLE"
                                                        , dataframe["ESTADO_SISCOVID"])
            
            dataframe["OBSERVACIÓN"] = np.where(
                ( (dataframe["Tipo_Documento"] == "DNI") & (dataframe["Nro_Documento"].str.len() != 8) ) |
                ( (dataframe["Tipo_Documento"] == "CE") & (dataframe["Nro_Documento"].str.len() != 9) )
                                                        , dataframe["OBSERVACIÓN"]+"Documento de identidad inválido para el número de dígitos."
                                                        , dataframe["OBSERVACIÓN"])
            
            
            
            #CASO DE INADMISIBLE : RESULTADO DE LA PRUEBA DIFERENTE A "REACTIVO" Y "NO REACTIVO"
            dataframe["ESTADO_SISCOVID"] = np.where(
                (dataframe["Resultado_de_la_prueba"] != "REACTIVO") &  (dataframe["Resultado_de_la_prueba"] != "NO REACTIVO")
                , "INADMISIBLE", dataframe["ESTADO_SISCOVID"])
            
            dataframe["OBSERVACIÓN"] = np.where(
                (dataframe["Resultado_de_la_prueba"] != "REACTIVO") &  (dataframe["Resultado_de_la_prueba"] != "NO REACTIVO")
                , dataframe["OBSERVACIÓN"] + "Resultado de prueba diferente de 'REACTIVO' y 'NO REACTIVO'.", 
                dataframe["OBSERVACIÓN"])
            

            #CASO INADMISIBLE: FECHA DE EJECUCIÓN DE LA PRUEBA ES MAYOR A HOY
            dataframe["ESTADO_SISCOVID"] = np.where(
                pd.to_datetime(dataframe["Fecha"], format="%d/%m/%Y") > datetime.now(),
                "INADMISIBLE", dataframe["ESTADO_SISCOVID"])
            
            dataframe["OBSERVACIÓN"] = np.where(
                pd.to_datetime(dataframe["Fecha"], format="%d/%m/%Y") > datetime.now(),
                dataframe["OBSERVACIÓN"] + "Fecha de toma de muestra aún no existe.", dataframe["OBSERVACIÓN"])
            
            
            
            #CASO FECHA DE EJECUCIÓN DE LA PRUEBA TIENE MÁS DE 40 DÍAS DE ANTIGUEDAD
            #SE CREA OBSERVACIÓN, SE MODIFICA FECHA DE EJECUCIÓN Y FECHA DE INICIO DE SÍNTOMAS
            
            
            dataframe["OBSERVACIÓN"] = np.where(
                pd.to_datetime(dataframe["Fecha"], format="%d/%m/%Y") <= datetime.now()- timedelta(days=40) ,
                    dataframe["OBSERVACIÓN"]+"TOMA DE MUESTRA REALIZADA EL "+dataframe['Fecha']+". REGULARIZACION DE REGISTRO POR ERROR DE SUBIDA EN SISTEMA."
                                                        , dataframe["OBSERVACIÓN"])
            
            dataframe["Fecha"] = np.where(
                pd.to_datetime(dataframe["Fecha"], format="%d/%m/%Y") <= datetime.now()- timedelta(days=40) ,
                (datetime.now()- timedelta(days=40)).strftime("%d/%m/%Y"), dataframe["Fecha"])
            
            dataframe["Fecha_de_inicio_de_Sintomas"] = np.where(
                (pd.to_datetime(dataframe["Fecha"], format="%d/%m/%Y") <= datetime.now()- timedelta(days=40)) & 
                (dataframe["Fecha_de_inicio_de_Sintomas"] != "") ,
                (datetime.now()- timedelta(days=40)).strftime("%d/%m/%Y"), dataframe["Fecha_de_inicio_de_Sintomas"])
            
        
            
            for index,row in dataframe.iterrows():
                
                #Corrige el DEPARTAMENTO, PROVINCIA Y DISTRITO en caso uno de ellos tenga un valor vacío ""
                if row["Departamento"] == "" or row["Provincia"] == "" or row["Distrito"] == "":
                    dataframe.loc[index,"Departamento"] = "LIMA"
                    dataframe.loc[index,"Provincia"] = "LIMA"
                    dataframe.loc[index,"Distrito"] = "LIMA"
                
                # Corrección Fecha de inicio de síntomas en caso tenga síntomas y la fecha de inicio de síntomas sea mayor a la fecha de ejecución de la prueba
                if row["Tiene_Sintomas"]=="SI":
                    if pd.to_datetime(dataframe.loc[index,"Fecha_de_inicio_de_Sintomas"], format="%d/%m/%Y") > pd.to_datetime(dataframe.loc[index,"Fecha"], format="%d/%m/%Y"):
                        dataframe.loc[index,"Fecha_de_inicio_de_Sintomas"] = dataframe.loc[index,"Fecha"]
                    
                    if dataframe.loc[index,"Marque_los_Sintomas_presenta"] == "":
                        dataframe.loc[index,"Marque_los_Sintomas_presenta"] = "TOS"
            
            #Guarda el archivo excel con los datos corregidos en la ruta especificada y se etiquera al índice
            dataframe.to_excel(self.save_path, index=True, index_label="ID")        
            return dataframe
        
        except Exception as e:
            logging.error("DataExcelSiscovid : Problema al ejecutar correccion_de_datos()")
            logging.error(f"{e.__class__}: {e}")
            raise FailDataExcelSiscovid(f"Error: {e.__class__}: {e}")
