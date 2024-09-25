import os
import time
from pathlib import Path
from platform import platform
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from dotenv import load_dotenv
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import sys
import pymysql
import pandas as pd
from datetime import datetime
from selenium.webdriver.chrome.options import Options as ChromeOptions

load_dotenv()
BASEDIR = Path('.').absolute()

class bot():

    def __init__(self):

        if 'Windows' in platform():
            self.user = os.getlogin()
            self.download_folder = os.path.join(f'C:\\Users\\prueba\\Downloads')
        else:
            # Para sistemas basados en Unix (Linux, macOS)
            self.download_folder = './Downloads'


        self.userWF = os.getenv('USER') 
        self.psw = os.getenv('PSW')
        self.files = []
        # self.files = ['Actividades-REGION OCCIDENTE_16_09_24']
        # Datos de conexión
        self.host = '190.60.100.100'       # Cambia esto por tu host
        self.user = 'BotCndCen'      # Cambia esto por tu usuario
        self.password = 'B0tCndC3n24*'  # Cambia esto por tu contraseña
        self.database = 'dbcrmcalv2'  # Cambia esto por tu base de datos
        self.port = 3306
        self.connection = self.get_db_connection()

    def get_db_connection(self):
        try:
            connection = pymysql.connect(
                host=self.host,
                user=self.user,
                password=self.password,
                database=self.database,
                port=self.port,
                charset='utf8mb4',
                cursorclass=pymysql.cursors.DictCursor
            )
            return connection
        except Exception as e:
            print(f'Error al conectar a la base de datos: {e}')
            return None

    def truncate_table(self):
        """Función para truncar la tabla"""
        try:
            with self.connection.cursor() as cursor:
                truncate_query = "TRUNCATE TABLE dbcrmcalv2.tbl_dataregioxhora;"
                cursor.execute(truncate_query)
                self.connection.commit()
                print("Tabla truncada correctamente.")
        except Exception as e:
            print(f"Error al truncar la tabla: {e}")

    def get_chrome_options(self):
        options = ChromeOptions()
        options.add_argument("--start-maximized")
        
        # Configurar las preferencias de descarga
        options.add_experimental_option("prefs", {
            "download.default_directory": '/home/seluser/Downloads',   # Directorio de descarga
            "download.prompt_for_download": False,  # No preguntar por la ubicación de descarga
            "download.directory_upgrade": True,  # Actualizar el directorio de descarga si cambia
            "safebrowsing.enabled": True,  # Habilitar la navegación segura
            "safebrowsing.disable_download_protection": True,  # Deshabilitar protección de descarga
            "profile.default_content_settings.popups": 0,  # Bloquear ventanas emergentes
            "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,  # Permitir descargas automáticas
            "profile.default_content_setting_values.automatic_downloads": 1,  # Permitir descargas múltiples
        })
        return options

    def create_browser(self):        

        if 'Windows' in platform():
            # print('The operating system is Windows\nWe will look for "Opera"')
            # from selenium.webdriver.opera.options import Options as OperaOptions

            # opera_options = OperaOptions()
            # opera_options.binary_location = r'%s\AppData\Local\Programs\Opera\opera.exe' % os.path.expanduser('~')
            # opera_options.add_argument('--start-maximized')
            # self.driver = webdriver.Opera(executable_path=r'C:\dchrome\operadriver.exe', options=opera_options)

            print('The operating system is Windows\nWe will look for "Chrome"')
            
            chrome_options = Options()
            chrome_options.add_argument('--start-maximized')  # Mantener otras opciones
            self.driver = webdriver.Chrome(executable_path=r'C:\dchrome\chromedriver.exe', options=chrome_options)

        else:
            print('The operating system is Linux\nWe will look for "Chrome"')
            time.sleep(10)
            chrome_options = self.get_chrome_options()
            self.driver = webdriver.Remote(
                command_executor='http://selenium:4444/wd/hub',
                options=chrome_options
            )

        # Abrir una nueva pestaña
        self.driver.execute_script("window.open('');")
        # Cambiar a la nueva pestaña
        self.driver.switch_to.window(self.driver.window_handles[1])
        self.driver.get('chrome://Downloads')
        time.sleep(1)
        self.driver.execute_script("window.open('');")
        self.driver.switch_to.window(self.driver.window_handles[2])
        self.driver.get('chrome://settings/downloads')
        time.sleep(1)
        self.driver.save_screenshot('./screenshot.png')
        self.driver.switch_to.window(self.driver.window_handles[0])

    def login_wf(self):
        time.sleep(10)
        self.driver.get("https://amx-res-co.etadirect.com/")                
        # Espera explícita para el campo de usuario
        wait = WebDriverWait(self.driver, 30)
        self.driver.save_screenshot('./screenshot.png')

        username_field = wait.until(EC.presence_of_element_located((By.ID, 'username')))
        
        # Limpiar y enviar el nombre de usuario
        username_field.clear()
        username_field.send_keys(self.userWF)
        self.driver.save_screenshot('./screenshot.png')
        
        # Espera explícita para el campo de contraseña
        password_field = wait.until(EC.presence_of_element_located((By.ID, 'password')))
        
        # Limpiar y enviar la contraseña
        password_field.clear()
        password_field.send_keys(self.psw)
        self.driver.save_screenshot('./screenshot.png')
        self.driver.execute_script('document.querySelector("#sign-in > div").click()')
        time.sleep(3)

        self.driver.save_screenshot('./screenshot.png')
    
        if self.driver.title=="Oracle Field Service":                
            bucle=1
            while bucle<=2:
                try:
                    self.driver.find_element(by=By.XPATH, value='//*[@id="username"]').clear()
                    self.driver.find_element(by=By.XPATH, value='//*[@id="username"]').send_keys(self.userWF)
                    self.driver.save_screenshot('./screenshot.png')
                    time.sleep(2)
                    self.driver.find_element(by=By.XPATH, value='//*[@id="password"]').clear()
                    self.driver.find_element(by=By.XPATH, value='//*[@id="password"]').send_keys(self.psw)
                    self.driver.save_screenshot('./screenshot.png')
                    time.sleep(1)
                    self.driver.find_element(by=By.XPATH, value='//*[@id="delsession"]').click()
                    time.sleep(2)
                    self.driver.execute_script('document.querySelector("#sign-in > div").click()')
                    time.sleep(3)
                    self.driver.save_screenshot('./screenshot.png')

                    break
                except Exception as e:
                    Nomb_error = 'Error on line {}'.format(sys.exc_info()[-1].tb_lineno), type(e).__name__, e
                    print("! error conexion: ", e, Nomb_error)
                    bucle+=1
            print(bucle)

            if bucle>=2:
                print('Error de login')
                self.driver.quit()        
            else:
                print("Login 2 vez ok!")
            
                try:
                    WebDriverWait(self.driver, 200).until(EC.invisibility_of_element_located((By.XPATH, '//div[@id="wait"]//div[@class="loading-animated-icon big jbf-init-loading-indicator"]')))
                except Exception as e:
                    print('')
                    self.create_browser()
                    self.login_wf()
                    self.search_tree()
                    exit
            
                time.sleep(1)            
                print(self.driver.title)
                if self.driver.title=="Cambiar contraseña - Oracle Field Service":
                    print('Cambiar contraseña - Oracle Field Service')
                    self.driver.save_screenshot('./screenshot.png')
                    self.driver.quit()
            self.driver.save_screenshot('./screenshot.png')

        else:
            try:
                WebDriverWait(self.driver, 200).until(EC.invisibility_of_element_located((By.XPATH, '//div[@id="wait"]//div[@class="loading-animated-icon big jbf-init-loading-indicator"]')))
            except Exception as e:
                print('')
                self.create_browser()
                self.login_wf()
                self.search_tree()
                exit            
            time.sleep(1)            
            print(self.driver.title)
            if self.driver.title=="Cambiar contraseña - Oracle Field Service":
                print('Cambiar contraseña - Oracle Field Service')
                self.driver.save_screenshot('./screenshot.png')
                self.driver.quit()

        self.driver.save_screenshot('./screenshot.png')

    def search_tree(self):

        vista_button = WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@title='Vista' and @aria-label='Vista']"))
        )
        # Haz clic en el botón "Vista"
        self.driver.save_screenshot('./screenshot.png')
        vista_button.click()

        checkbox = WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[26]/div/div/div/div[1]/form/div/div[3]/oj-checkboxset'))
        )
        self.driver.save_screenshot('./screenshot.png')
        checkbox.click()

        # Encuentra el contenedor principal
        container = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[26]/div/div/div/div[2]'))
        )

        button_html = container.get_attribute('outerHTML')
        print("HTML del botón 'Aplicar':\n", button_html)
        self.driver.save_screenshot('./screenshot.png')
        container.click()

        if 'Windows' in platform():
            button = container.find_element_by_css_selector(".app-button--cta")  # Target the 'Aplicar' button
        else:
            button = container.find_element(By.CSS_SELECTOR, ".app-button--cta")  # Target the 'Aplicar' button
            
        self.driver.save_screenshot('./screenshot.png')
        button.click()
        print("Successfully clicked the 'Aplicar' button!")

        container = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="manage-content"]/div/div[2]/div[3]/div/div[1]/table/tbody/tr[2]/td[1]/div/div/div[1]'))
        )

        list = WebDriverWait(self.driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="manage-content"]/div/div[2]/div[2]/div/div[2]/div[3]'))
        )
        
        lista_regiones = [
            "REGION OCCIDENTE",
            "PYMES OCCIDENTE",
            "DTH OCCIDENTE",
            "R5 ORIENTE",
            "R4-Tabasco CENTRO"
        ]
        cont = 0

        for i in [8,15,17,18,20]:
            try:
                path = f'//*[@id="manage-content"]/div/div[2]/div[2]/div/div[2]/div[3]/div[{i}]/div[1]/button[3]/span[1]'
                apply_button = WebDriverWait(list, 10).until(
                    EC.presence_of_element_located((By.XPATH, path))
                )
                region_text = apply_button.text
                self.driver.save_screenshot('./screenshot.png')

                if region_text in lista_regiones:
                    print('el numero: ',i)
                    print('')
                    print('clik a: ',region_text)
                    apply_button.click()
                    wait_title = 0

                    while wait_title < 5:
                        wait_title +=1
                        titulo = WebDriverWait(self.driver, 10).until(
                            EC.visibility_of_element_located((By.XPATH, '//div[@class="page-header-description page-header-description--text"]'))
                        )

                        title_text = titulo.get_attribute('innerText')
                        print(title_text)
                        # print(region_text)
                        if title_text == region_text:
                            print('ir a descargar')

                            while True:
                                acciones_button = WebDriverWait(self.driver, 30).until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[@title='Acciones' and @aria-label='Acciones']"))
                                )
                                self.driver.save_screenshot('./screenshot.png')
                                acciones_button.click()

                                exportar = WebDriverWait(self.driver, 30).until(
                                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[26]/div/div/button[2]'))
                                )
                                self.driver.save_screenshot('./screenshot.png')
                                exportar.click()
                                
                                ttt = 0
                                while not ttt>5:
                                    ttt += 1
                                    self.driver.save_screenshot('./screenshot.png')
                                    time.sleep(1)

                                # ttt = 0
                                while not self.archivos_descargados(region_text):
                                # while not ttt>5:
                                    print('entra a buscar...')
                                    # ttt += 1
                                    # self.archivos_descargados(region_text)

                                    self.driver.switch_to.window(self.driver.window_handles[1])
                                    self.driver.save_screenshot('./screenshot.png')
                                    time.sleep(3)
                                    self.driver.switch_to.window(self.driver.window_handles[0])
                                    self.driver.save_screenshot('./screenshot.png')
                                    time.sleep(3)
                                self.driver.switch_to.window(self.driver.window_handles[0])
                                print(self.files)
                                time.sleep(3)
                                break
                            break
                        time.sleep(2)

                    cont += 1
                    if cont == len(lista_regiones):
                        break

            except Exception as e:
                print(e)
                break

        print('listo.!')
        self.driver.quit()        
        self.upload_data()

    def archivos_descargados(self, region):

        archivos_presentes = os.listdir(self.download_folder)
        encontrado = False
        # print(archivos_presentes)

        for archivo in archivos_presentes:
            print('Buscando archivo: ', archivo)
            archivo_base, extension = os.path.splitext(archivo)
            print(archivo_base)
            if extension == '.opdownload':
                archivo = archivo_base  # Elimina la extensión

            if archivo.startswith(f"Actividades-{region}") and archivo.endswith(".xlsx"):
                encontrado = True
                self.files.append(archivo)
                print(f"Archivo encontrado: {archivo}")
                break

        if not encontrado:
            return False
        return True

    def upload_data(self):
        print('inicia a cargar...')
        try:

            with self.connection.cursor() as cursor:
                successful = 0
                error = 0
                num_registros = 0
                total_registros = 0
                try:

                    for data_file in self.files:
                        csv_file = os.path.join(self.download_folder, data_file)
                        print('el archivo: ', csv_file)
                        df = pd.read_excel(csv_file, engine='openpyxl')
                        csv_file = os.path.join(self.download_folder, data_file)

                        # Reemplaza NaN con una cadena vacía
                        df = df.fillna('')
                        # Convierte todas las columnas a str
                        df = df.astype(str)

                        # print(df.columns)
                        # print(list(df.columns))
                        successful = 0
                        error = 0
                        num_registros = 0

                        # Obtén el número de registros (filas) sin contar el encabezado
                        num_registros = df.shape[0] - 1
                        self.truncate_table()

                        for _, row in df.iterrows():
                            try:
                                # Prepara los datos del registro
                                params = (
                                    row['Técnico'], row['Intervalos de tiempo'], row['ID Aliado'], row['Fecha'], row['Nombre'],row['Código Asesor comercial'], row['Dirección campo 1'], row['Nombre Completo'], 
                                    row['Tipo de Actividad'],row['Subtipo de la Orden de Trabajo'], row['Orden de trabajo Mantenimientos 7K'], row['Orden de Trabajo'],row['Estado'], row['Zonas de trabajo'], 
                                    row['Aptitud laboral'], row['Zona'], row['Inicio'],row['Fin'], row['Inicio - Fin'], row['Inicio de SLA'], row['Fin de SLA'], row['Tiempo de viaje'],row['Ventana de servicio'], 
                                    row['Ventana de entrega'], row['Duración'], row['Ciudad'], row['Departamento '],row['Nodo'], row['Número de cuenta'], row['Estado interno de la OT'], row['Numero de Reincidencias Serivicios'],
                                    row['Numero de Reincidencias Calidad'], row['Numero de Reprogramaciones'], row['SLA Suscriptor'],row['SLA Cumplimiento'], row['Estado SLA'], row['Asesor comercial'], row['Tipo de Red'],
                                    row['Materials Validation Result'], row['Resultados SoftClose'], row['Regional'], row['Unidad de Gestión'],row['Razón'], row['Número de Ticket Fallas Masivas'], row['Código de causa 1 del IIMS'],
                                    row['Código de causa 2 del IIMS'], row['Código de causa 3 del IIMS'], row['Fecha de agendamiento'],row['Compañia'], row['External ID'], row['Persona que Confirma'], row['Adjunto Evidencia'], 
                                    row['Caso SD'],row['Dirección E-mail Solicitante'], row['Nombre Completo'], row['Tipo Solucion'], row['Fecha de gestion'],row['No. OT Backoffice'], row['Cuenta Cliente'], row['Nombre Cliente'], 
                                    row['Dirección '],row['Tipo Cliente'], row['Tipo OT'], row['Ciudad'], row['Region'], row['ID Nodo'], row['Aliado'],row['Aliado CGO'], row['Tipo Gestion'], row['Causa Gestión'], row['Detalle Gestión'], 
                                    row['Resultado Gestión'],row['Notas Resultado'], row['No. Aviso'], row['TIPO GESTION BACKLOG'], row['ALIADO CGO BACLOG'],row['Primera operación manual realizada por usuario'], 
                                    row['Primera operación manual realizada por usuario (conexión)'],row['Primera operación manual realizada por usuario (nombre)'], row['BACKLOG_GESTION'], row['BACKLOG_NOMBRE'],
                                    row['BACKLOG_PARENTESCO'], row['BACKLOG_MEDIOCONTACTO'], row['CONFIRMACION'], row['CONFIR_RESULTADO'],row['CONFIR_MEDIO'], row['CONFIR_USUARIOCO'], row['CONFIR_ALIADOCGO'], 
                                    row['CCICLO_ALIADOCGO'],row['CCICLO_PERSONAENC'], row['CGE_ALIADOCGO'], row['EXPERTO_CAUSA'], row['EXPERTO_GESTION'],row['ESCALAM_CAUSA'], row['ESCALAM_GESTION'], row['ESCALAM_AVISO'], 
                                    row['DESPAC_ALIADOCGO'],row['DEPACH_REPORTEDEMO'], row['DEPACH_HORAREPORTE'], row['DESPAC_CAUSA'], row['DESPAC_GESTION'],row['CAMBIOEQ_CAUSA'], row['CAMBIOEQ_GESTION'], row['ENRUTAM_ALIADO'], 
                                    row['ENRUTAM_MARCA'],row['Prioridad'], row['Parent account'], row['Agenda Inmediata'], row['Número Marker'],row['Contacto Teléfono 2'], row['Telefono Contacto'], row['Telefono Encuesta'],
                                    row['Telefono dos del cliente'], row['Teléfono 3'], row['Teléfono uno del contacto'],row['Lista Reporte Demora'], row['Activacion Automatica'], row['Num Reprogramaciones'],row['Causa solicitud TAM'], 
                                    row['Gestion TAM'], row['Aliado'], row['INSTALACION ANDROID TV'],row['Razones 7K'], row['Id Sitio CD'], row['Soporte Proyectos'], row['Aliado CND Proyecto'],row['CND que Gestiona'], 
                                    row['Causa Escalamiento Proyecto'], row['Tiempo duracion Soporte'],row['Ubicación en la ruta'], row['Usuario'], row['Aliado CGO Confirmación'], row['Minuto'],row['Tipo servicio'], 
                                    row['Ot Modulo Gestion'], row['Regional N&E'], row['No OT Hija N&E Man'],row['Razón de NO Realización '], row['Contacto de la orden de trabajo'], row['Contacto del edificio'],row['Closure Notes'], 
                                    row['Numero OS/oth'], row['Número OP/otp'], row['Estado OP/otp'],row['Estado OS/oth'], row['Estado OS/oth'], row['Nombre del edificio'], row['Direccion Alterna'],row['Dirección del edificio'], 
                                    row['OT Padre N&E Ins'], row['Segmento'], row['SubTipo de Trabajo'],row['OT Modulo de Gestion '], row['Cedula asesor que atiende'], row['CONFIR_CHECK'], row['Grupo de Actividades'],row['Actividades'], 
                                    row['Numero de OT proyecto'], row['Subtipo actividad'], row['ID troncal o Nodo'],row['Realizo actividad?'], row['Documento actividad?'], row['Activo GPS?'], row['Fecha y Hora de inicio'],
                                    row['Fecha y hora de finalización'], row['Inicio actividad'], row['Fecha creación CRM-Dia'], row['Mes'],row['Año'], row['Minuto'], row['Hora creacion CRM'], row['Hora'], row['No OT Padre N&E Man'],
                                    row['Prioridad'], row['Hora'], row['Regional'], row['Numero Cedula de quien confirma'], row['Segmento'],row['Tarea Mantenimiento FO'], row['ID OT RR'], int(row['Actividad ID'])
                                )
                                print('3')
                                
                                sql = """
                                CALL spr_dataloadingGes(
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s
                                )
                                """
                                print('3')
                                
                                self.clear_console()
                                print('3')

                                # Imprime los nuevos datos
                                print('')
                                print(data_file)
                                print(num_registros, ' - ' ,successful)
                                if successful != 0:
                                    porcentaje = (successful / num_registros) * 100
                                    print(f"{porcentaje:.2f}%")

                                if error >0:
                                    print('errores: ',error)

                                print(params)
                                # print(f"\rData file: {data_file}, Params: {params}, Successful: {successful}", end=NULL, flush=True)
                                
                                cursor.execute(sql, params)
                                successful +=1
                            except Exception as e:
                                error += 1
                                print(f"Error al intentar insertar el valor truncado: {e}")

                        # Commit de la transacción
                        self.connection.commit()
                        print("Datos insertados correctamente.")
                        os.remove(csv_file)
                    print('finalizo archivo')

                except Exception as e:
                    print('Error al leer archivos: ', e)

                print('de ',data_file,' se cargaron ', successful, ' de ', num_registros, ' registros' )
                total_registros += num_registros
                print('Total registros ',total_registros)
                successful, error = 0

        except Exception as e:
            print(f'Error al conectar a la base de datos: {e}')

        self.connection.close()

    def clear_console(self):
        if os.name == 'nt':  # Para Windows
            os.system('cls')
        else:  # Para Linux y macOS
            os.system('clear')

    def win(self):


        self.driver.get("https://winrar.es/descargas")                
        self.driver.save_screenshot('./screenshot.png')
        
        # Espera explícita para el campo de usuario
        wait = WebDriverWait(self.driver, 30)
        self.driver.save_screenshot('./screenshot.png')

        button = WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="dwrecbox"]'))
        )
        self.driver.save_screenshot('./screenshot.png')

        button.click()
        self.driver.save_screenshot('./screenshot.png')


        ttt = 0
        while not ttt>10:
            ttt += 1
            print(ttt)
            self.driver.save_screenshot('./screenshot.png')
            time.sleep(1)


        while not ttt>5:
            ttt += 1
            self.driver.switch_to.window(self.driver.window_handles[1])
            self.driver.save_screenshot('./screenshot.png')
            time.sleep(3)
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.driver.save_screenshot('./screenshot.png')
            time.sleep(3)

if __name__ == '__main__':

    # if 'Windows' in platform():

    #     while True:
    #         now = datetime.now()
    #         if now.minute >= 20 and  5 < now.hour < 22:
    #             print("Es hora de imprimir: ", now.strftime("%H:%M:%S"))
    #             b = bot()
    #             b.create_browser()
    #             b.login_wf()
    #             b.search_tree()

    #         time.sleep(10)

    b = bot()
    b.create_browser()
    
    # b.win()
    
    b.login_wf()
    b.search_tree()

    # b.upload_data()
