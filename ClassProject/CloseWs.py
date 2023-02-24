import os
import win32com.client
import logging

class FailCloseWs(Exception):
    """Excepcion que se lanza cuando no se puede cerrar excel o word al
    ejecutar la clase CloseWs"""
    def __init__(self, msg):
        self.msg=msg
    
    def __str__(self):
        return self.msg

class CloseWs:
    
    @staticmethod
    def cerrar_excel_y_word() -> bool:
        """Funcion que termina el proceso del programa excel y word en caso de estar abierto.
        
        Raises:
            failCloseWs: Error en caso de que la función no pueda ser ejecutada de manera correcta
        
        Returns:
            bool: True - Excel y Word ya no se ejecutan      
        """
        
        try:
            CloseWs.close_excel()
            CloseWs.close_word()
            return True
        except Exception:
            logging.error("CloseWs : Problema al ejecutar cerrar_excel_y_word()")
            raise FailCloseWs("No se pudo cerrar excel y word")
            
    @staticmethod
    def close_excel() -> bool:
        """Funcion que termina el proceso del programa excel en caso de estar abierto.

        Raises:
            FailCloseWs: Error en caso de que la función no pueda ser ejecutada de manera correcta

        Returns:
            bool: True - Excel ya no se ejecuta
        """
        try:
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                excel.Quit()
                os.system("taskkill /f /im excel.exe")
                return True
            except Exception:
                return True
        except Exception:
            logging.error("CloseWs : Problema al ejecutar close_excel()")
            raise FailCloseWs("No se pudo cerrar excel")

    @staticmethod
    def close_word() -> bool:
        """Funcion que termina el proceso del programa word en caso de estar abierto.
        
        Raises:
            FailCloseWs: Error en caso de que la función no pueda ser ejecutada de manera correcta
        
        Returns:
            bool: True - Word ya no se ejecuta
        """
        try:         
            try:
                word = win32com.client.GetActiveObject("Word.Application")
                docs = word.Documents
                for doc in docs:
                    doc.Close(False)
                word.Quit()
                os.system("taskkill /f /im winword.exe")
                return True
                
            except Exception:
                return True
        
        except Exception:
            logging.error("CloseWs : Problema al ejecutar close_word()")
            raise FailCloseWs("No se pudo cerrar word")
            