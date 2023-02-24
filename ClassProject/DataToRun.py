from ClassProject.DataExcel import DataExcel
from ClassProject.CloseWs import CloseWs
import os
import win32com.client

class FailDataApp(Exception):
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

class DataApp(DataExcel):
   
    def __init__(self):
        #Sobre escribe el atributo path para que sea de la clase DataApp
        self.path=self.find_data_path()
        #Inicializa la clase padre
        super().__init__(path=self.path, header="Fecha", dtype_str=True)
        self._frame=self.load_app_data()
    
    @property
    def frame(self):
        return self._frame
    
    def find_data_path(self):
        data_path=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                              "Data","Data_to_run.xlsx")
        return data_path
    
    def load_app_data(self):
        CloseWs.close_excel()
        dataframe=self.load_excel()
        dataframe.set_index("ID", inplace=True,drop=True)
        dataframe.fillna("", inplace=True)
        dataframe=dataframe.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        return dataframe
    
   
    def open_excel(self):
        CloseWs.close_excel()
        # Crea un objeto Excel
        excel = win32com.client.Dispatch("Excel.Application")
        # Muestra la ventana de Excel
        excel.Visible = True
        # Abre el archivo especificado
        excel.Workbooks.Open(self.path)
    
    
    
        
        
        