import os
import datetime
from os import listdir
from os.path import isfile, isdir
import pandas as pd
import xlrd
from tqdm import tqdm
import pandas as pd
from pandas import ExcelWriter



class Dataplan(object):
    """
    Dataplan class
    Crea un "Dataplan de microwaves" a partir de archivos de "Microwaves Configuration".

    Parameters
    ----------
    path : string
        Es la ruta de donde se encuentan los archivos para procesar
        solo se tomaran en cuenta los archivos con extencion ".xls" o ".xlsx"
        en caso que este se encuentre vacio, tomara la ruta donde se esta trabajando.
    name : string
        Es el nombre que le vamos asignar a nuestro "Dataplan" generado, tomar en cuenta
        que el nombre de nuestro archivo resultante contiene adicionalmente la fecha y hora
        de la creacion

    """

    def __init__(self, path,nombre):
        super(Dataplan, self).__init__()
        self.default_name = "Dataplan"
        self.sheet = "Sheet0"
        self.default_path = os.getcwd()+"/"
        self.path_files = path
        self.name = nombre

    def ls(self):
        """
        Lista los archivos de la ruta o "path" que hemos asignado.

        Returns
        -------
        List
            Devuelve una lista de archivos encontrados en la ruta "path"

        Raises
        ------
        KeyError
            empy list
        """
        path = ""
        if self.path_files != "":
            path = self.path_files
        else:
            path = self.default_path
        return [obj for obj in listdir(path) if isfile(path + obj) and ".~" not in obj[:2] and ".xls" in obj[obj.find("."):]]

    def get_dataplan(self,console=False):
        """
        Crea el archivo ".xlsx" con el nombre asignado.

        Parameters
        ----------
        console : Boolean
            Se utliza cuando el entorno de trabajo es una consola

        Returns
        -------
        List
            Error[index]: Error en un archivo especifico, se listan la cantidad
                          de errores que existan en el proceso.

            errors : La cantidad de errores en el proceso

            status : El mensaje para el usuario

            file : Ruta y nombre del archivo creado

        Raises
        ------
        KeyError
            Mensaje dento de la lista que regresa
        """
        messages = []
        path_files = ""
        if self.path_files != "":
            path_files = self.path_files
        else:
            path_files = self.default_path
        files = self.ls()
        validate_console = files
        c = len(files)
        if console:
            validate_console = tqdm(files)
            for file in files:
                print(file)
            print(f"\nTotal: {c} Archivos")

        microwaves = []

        if c != 0:
            archivos_malos = 0
            for file in validate_console:
                except_this = ""
                nombre = file[:file.find("_")]

                #Ruta del archivo
                archivo = path_files+file
                #abrir Archivo
                try:
                    libro = xlrd.open_workbook(archivo)
                    sheet = libro.sheet_by_name(self.sheet)
                except Exception as e:
                    if console:
                        print(f"El Archivo {file} no es compatible")
                    except_this = "No compatible"
                    archivos_malos += 1
                    messages.append({
                    "Error["+str(archivos_malos)+"]": file+" no es compatible"
                    })
                eth_type = ''
                source_node = ''
                ports_eth = ''
                if except_this == "":
                    for i in range(sheet.nrows):
                        if "ETH: E-Line" in sheet.cell_value(i,0):
                            eth_type = sheet.cell_value(i,0)
                            j = 2
                            source_node = ''
                            while sheet.cell_value(i+j,3) != '':
                                port = sheet.cell_value(i+j,3)
                                add = port[:port.find("[")]+'\n'
                                if add not in source_node:
                                    source_node += add
                                j += 1
                            service = eth_type[5:]


                        elif "ETH: E-LAN" in sheet.cell_value(i,0):
                            eth_type = sheet.cell_value(i,0)
                            j = 6
                            source_node = ''
                            while sheet.cell_value(i+j,3) != '' and "Port Enable" not in sheet.cell_value(i+j,3):
                                port = sheet.cell_value(i+j,3)
                                add = port[:port.find("[")]+'\n'
                                if add not in source_node:
                                    source_node += add
                                j += 1
                            service = eth_type[5:]

                        elif "Port Information for ETH" in sheet.cell_value(i,0):
                            f = 2
                            while sheet.cell_value(i+f,3) != '' and i+f+1 != sheet.nrows:
                                if sheet.cell_value(i+f,3) == 'Enabled' and sheet.cell_value(i+f,1) not in ports_eth:
                                    ports_eth += sheet.cell_value(i+f,1)+'\n'
                                f += 1

                        elif "NE Type:" in sheet.cell_value(i,0):
                            value = sheet.cell_value(i,0)
                            RTN_equipment = value[8:]

                        elif "SDH/PDH Service" in sheet.cell_value(i,0):
                            eth_type = sheet.cell_value(i,0)
                            j = 2
                            while sheet.cell_value(i+j,3) != '':
                                port = sheet.cell_value(i+j,4)
                                source_node += port[:port.find("[")]+'\n'
                                j += 1
                                service = eth_type
                    microwaves.append({
                        "ID":nombre,
                        "RTN Equipment":RTN_equipment,
                        "Version":'',
                        "SUBNET":'',
                        "Ports ETH":ports_eth,
                        "Service":service,
                        "Ports IF":source_node,
                    })

            archivo_final = ""
            if self.name != "":
                archivo_final = self.name
            else:
                archivo_final = self.default_name
            fecha = str(datetime.datetime.now())
            if archivos_malos != len(files):
                if archivos_malos == 0:
                    messages.append({
                    "status":"Todos los Archivos procesados con exito",
                    "file": os.getcwd()+"/"+archivo_final+"_"+fecha+'.xlsx',
                    "errors":archivos_malos
                    })
                else:
                    messages.append({
                    "status":str(archivos_malos)+" de "+str(len(files))+" no fueron procesados",
                    "file": os.getcwd()+"/"+archivo_final+"_"+fecha+'.xlsx',
                    "errors":archivos_malos
                    })
                df = pd.DataFrame(microwaves)
                df = df[['ID','RTN Equipment','Version','SUBNET','Ports ETH', 'Service', 'Ports IF']]
                writer = ExcelWriter(os.getcwd()+"/"+archivo_final+"_"+fecha+'.xlsx')
                df.to_excel(writer, archivo_final, index=False)
                writer.save()
                if console:
                    print(f"\nSe ha creado el archivo {archivo_final}_{fecha}.xlsx en {os.getcwd()}")
            else:
                messages.append({
                "status":"Todos los archivos son incompatibles",
                "file":"",
                "errors":archivos_malos
                })



        else:
            if console:
                print("No se encontraron archivos para procesar")
            messages.append({
            "status":"No se encontraron archivos para procesar",
            "file":"",
            "errors":archivos_malos
            })
        return messages
