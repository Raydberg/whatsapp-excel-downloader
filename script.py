from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import os

# =================== CONFIGURACIÓN PRINCIPAL ===================

# Configuración del contacto y chat
NOMBRE_CONTACTO = "Darling" 

# Configuración de Chrome
RUTA_CHROME = "C:/Program Files/Google/Chrome/Application/chrome.exe"
PERFIL_CHROME = "PERSONAL"
CARPETA_DESCARGAS = "E:/python/DOCUMENTOS"

# Configuración de tiempos de espera (segundos)
TIEMPO_ESCANEO_QR = 15
TIEMPO_CARGA_PAGINA = 20
TIEMPO_ENTRE_ACCIONES = 1.5
TIEMPO_DESCARGA = 3

# Configuración de búsqueda de archivos
EXTENSIONES_PERMITIDAS = [".xlsx", ".xls", ".csv"]
MAX_SCROLLS = 10
MAX_SCROLLS_SIN_RESULTADOS = 3

# 🔹 Configurar Chrome
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = RUTA_CHROME
chrome_options.add_argument(f"--profile-directory={PERFIL_CHROME}")

# Configurar carpeta de descargas
prefs = {"download.default_directory": CARPETA_DESCARGAS}
chrome_options.add_experimental_option("prefs", prefs)
archivos_procesados = set()

# 🔹 Iniciar WebDriver
driver = webdriver.Chrome(options=chrome_options) 

# Abrir WhatsApp Web
driver.get("https://web.whatsapp.com")
print("🔄 Esperando que escanees el código QR...")
time.sleep(TIEMPO_ESCANEO_QR)

#=================== FUNCIONES ===================

def abrir_chat(contacto):
    """
    Función mejorada para abrir un chat usando un enfoque más directo
    con simulación de comportamiento humano
    """
    print(f"🔍 Buscando al contacto: {contacto}...")
    
    # Dar tiempo para cargar la página completamente
    time.sleep(TIEMPO_CARGA_PAGINA)
    
    # Captura del estado inicial
    
    try:
        # Abrir directamente el chat usando la URL con teléfono/nombre codificado
        chat_url = f"https://web.whatsapp.com/search/{contacto}"
        driver.get(chat_url)
        print(f"✓ Navegando directamente a la búsqueda: {chat_url}")
        time.sleep(5)  # Esperar a que cargue la búsqueda
        
        
        # Buscar el contacto en los resultados de búsqueda
        contact_selectors = [
            f'//span[@title="{contacto}"]',
            f'//span[contains(@title, "{contacto}")]',
            f'//div[contains(@title, "{contacto}")]'
        ]
        
        for selector in contact_selectors:
            try:
                chat_element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                print(f"✓ Contacto encontrado con selector: {selector}")
                
                # Simular comportamiento humano: scroll + espera + clic
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", chat_element)
                time.sleep(TIEMPO_ENTRE_ACCIONES)
                
                # Usar ENTER después de hacer clic (esto a veces funciona mejor)
                chat_element.click()
                chat_element.send_keys(Keys.ENTER)
                time.sleep(4)
                
                
                # Esperar explícitamente a que cambie la URL o la UI
                WebDriverWait(driver, 10).until(lambda d: "search" not in d.current_url)
                print("✅ Navegación a chat detectada")
                
                time.sleep(3)  # Dar tiempo para que termine de cargar
                break
            except Exception as e:
                print(f"Intento fallido con selector {selector}: {str(e)}")
        
        verificaciones = [
            ('//header//div[contains(@data-testid, "conversation-info")]', "Encabezado de conversación"),
            ('//footer//div[@role="textbox"]', "Campo de mensaje"),
            ('//div[contains(@data-testid, "conversation-panel")]', "Panel de conversación")
        ]
        
        for selector, descripcion in verificaciones:
            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, selector)))
                print(f"✅ Verificación exitosa: {descripcion} encontrado")
                return True
            except:
                continue
        
        if "search" not in driver.current_url:
            print("✅ Verificación por URL: Chat abierto correctamente")
            return True
        
        
        try:
            driver.execute_script(f"""
                document.querySelectorAll('div[role="button"]').forEach(function(el) {{
                    if (el.innerText.includes("{contacto}")) {{
                        el.click();
                    }}
                }});
            """)
            time.sleep(5)
            return True
        except:
            pass
            
        raise Exception("No se pudo confirmar la apertura del chat")
        
    except Exception as e:
        print(f"❌ Error al abrir chat: {str(e)}")
        raise Exception(f"No se pudo abrir el chat de {contacto}: {str(e)}")

def extraer_nombre_archivo(texto):
    """Extrae el nombre real del archivo desde diferentes formatos de texto"""
    if not texto:
        return ""
    
    # Caso 1: El texto es solo un nombre de archivo
    if any(texto.endswith(ext) for ext in EXTENSIONES_PERMITIDAS):
        return texto
    
    # Caso 2: El texto incluye detalles después del nombre
    lines = texto.split('\n')
    if lines and any(lines[0].endswith(ext) for ext in EXTENSIONES_PERMITIDAS):
        return lines[0]
    
    # Caso 3: Intentar extraer patrón de nombre de archivo con extensión
    import re
    extensions_pattern = '|'.join(ext.replace('.', '\\.') for ext in EXTENSIONES_PERMITIDAS)
    match = re.search(f'([^\\s]+\\.({extensions_pattern[1:]}))', texto)
    if match:
        return match.group(1)
    
    # Si no podemos determinar el nombre, devolver el texto original
    return texto

def verificar_archivo_existe(nombre_archivo):
    """Verifica si un archivo ya existe en la carpeta de descargas"""
    if not nombre_archivo:
        return False
        
    # Si el nombre no tiene extensión válida, no podemos verificarlo correctamente
    if not any(nombre_archivo.endswith(ext) for ext in EXTENSIONES_PERMITIDAS):
        return False
    
    # Verificar si existe en la carpeta de descargas
    ruta_completa = os.path.join(CARPETA_DESCARGAS, nombre_archivo)
    return os.path.exists(ruta_completa)

def buscar_y_descargar_archivos_en_vista_actual(max_intentos=None):
    """Busca y descarga archivos Excel que están actualmente visibles en la pantalla"""
    try:
        print("🔍 Buscando archivos recientes en la vista actual...")
        global archivos_procesados  # Usar la variable global
        
        # Detección de mensajes con selectores alternativos
        message_selectors = [
            "//div[contains(@class, 'message-in') or contains(@class, 'message-out')]",
            "//div[@data-testid='msg-container']",
            "//div[contains(@class, 'focusable-list-item')]",
            "//div[contains(@class, 'message')]"
        ]
        
        mensajes = []
        for selector in message_selectors:
            try:
                found_messages = driver.find_elements(By.XPATH, selector)
                if found_messages:
                    # Invertir la lista para procesar primero los más recientes
                    mensajes = found_messages[::-1]
                    print(f"✅ Encontrados {len(mensajes)} mensajes con selector: {selector}")
                    break
            except:
                continue
        
        if not mensajes:
            print("⚠️ No se encontraron mensajes en la vista actual")
            return 0
        
        # Estadísticas de procesamiento
        document_count = 0
        excel_count = 0
        downloads_attempted = 0
        
        for mensaje in mensajes:
            # Si se estableció un límite y ya lo alcanzamos, salimos
            if max_intentos is not None and downloads_attempted >= max_intentos:
                print(f"✅ Se alcanzó el límite de {max_intentos} archivos descargados")
                break
                
            try:
                message_id = mensaje.get_attribute("data-id") or mensaje.get_attribute("id") or ""
                if not message_id:
                    message_text = mensaje.text[:50] if mensaje.text else ""
                    rect = driver.execute_script("return arguments[0].getBoundingClientRect();", mensaje)
                    message_id = f"{message_text}_{rect['top']}_{rect['left']}"
                
                if message_id in archivos_procesados:
                    continue
                
                document_selectors = [
                    ".//div[contains(@data-testid, 'document')]",
                    ".//div[contains(@data-testid, 'document-thumb')]",
                    ".//div[contains(@aria-label, 'document')]",
                    ".//a[contains(@href, 'blob:')]",
                    ".//img[@data-testid='document-thumb']",
                    f".//span[contains(text(), '{EXTENSIONES_PERMITIDAS[0]}') or contains(text(), '{EXTENSIONES_PERMITIDAS[1]}') or contains(text(), '{EXTENSIONES_PERMITIDAS[2]}')]",
                    f".//div[contains(@role, 'button') and .//span[contains(text(), '{EXTENSIONES_PERMITIDAS[0]}') or contains(text(), '{EXTENSIONES_PERMITIDAS[1]}') or contains(text(), '{EXTENSIONES_PERMITIDAS[2]}')]]",
                    ".//div[contains(@class, 'document-message-container')]"
                ]
                
                for selector in document_selectors:
                    posibles_archivos = mensaje.find_elements(By.XPATH, selector)
                    document_count += len(posibles_archivos)
                    
                    for archivo in posibles_archivos:
                        if max_intentos is not None and downloads_attempted >= max_intentos:
                            break
                            
                        texto = archivo.text.lower() if archivo.text else ""
                        href = archivo.get_attribute("href") or ""
                        aria_label = archivo.get_attribute("aria-label") or ""
                        data_testid = archivo.get_attribute("data-testid") or ""
                        class_name = archivo.get_attribute("class") or ""
                        
                        nombre_archivo = extraer_nombre_archivo(texto)
                        print(f"📄 Nombre de archivo extraído: '{nombre_archivo}'")
                        
                        if nombre_archivo and verificar_archivo_existe(nombre_archivo):
                            print(f"⏭️ Archivo '{nombre_archivo}' ya existe en la carpeta de descargas, saltando...")
                            archivos_procesados.add(f"{nombre_archivo}")
                            archivos_procesados.add(message_id)
                            continue
                            
                        file_id = f"{nombre_archivo}" if nombre_archivo else f"{texto}_{href}_{aria_label}_{data_testid}"
                        
                        if file_id in archivos_procesados:
                            print(f"⏭️ Archivo ya procesado anteriormente: {nombre_archivo or texto}")
                            continue
                        
                        archivos_procesados.add(file_id)
                        archivos_procesados.add(message_id)
                        
                        print(f"🔎 Elemento encontrado: texto='{texto}', data-testid='{data_testid}', aria-label='{aria_label}'")
                        
                        extensions = [ext[1:] for ext in EXTENSIONES_PERMITIDAS] + ["excel", "reporte"]
                        es_excel = any(ext in texto.lower() for ext in extensions) or \
                                  any(ext in href.lower() for ext in extensions) or \
                                  any(ext in aria_label.lower() for ext in extensions) or \
                                  ("document" in data_testid) or \
                                  (nombre_archivo and any(nombre_archivo.endswith(ext) for ext in EXTENSIONES_PERMITIDAS))
                        
                        if es_excel:
                            excel_count += 1
                            print(f"📊 Archivo Excel encontrado: {nombre_archivo or texto}")
                            
                            # Hacer visible y descargar
                            try:
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", archivo)
                                time.sleep(TIEMPO_ENTRE_ACCIONES)
                                
                                archivo.click()
                                time.sleep(1)
                                
                                download_buttons = driver.find_elements(By.XPATH, 
                                    "//div[@role='button' and contains(@aria-label, 'Descargar') or contains(@aria-label, 'Download')]")
                                
                                if download_buttons:
                                    download_buttons[0].click()
                                    downloads_attempted += 1
                                    print(f"✅ Descarga #{downloads_attempted} iniciada: {nombre_archivo or texto}")
                                else:
                                    try:
                                        driver.execute_script("arguments[0].click();", archivo)
                                        time.sleep(1)
                                        downloads_attempted += 1
                                        print(f"✅ Descarga #{downloads_attempted} iniciada (método JS): {nombre_archivo or texto}")
                                    except:
                                        print("⚠️ No se pudo descargar con método alternativo")
                                
                                time.sleep(TIEMPO_DESCARGA)  
                            except Exception as e_click:
                                print(f"⚠️ No se pudo descargar: {str(e_click)}")
                                
            except Exception as e:
                print(f"⚠️ Error procesando mensaje: {str(e)}")
                continue
                
        print(f"📑 Análisis completo: {document_count} documentos encontrados, {excel_count} archivos Excel, {downloads_attempted} intentos de descarga")
        return downloads_attempted
        
    except Exception as e:
        print(f"⚠️ Error al buscar archivos: {str(e)}")
        return 0

def buscar_descargar_excel_existentes():
    """Busca y descarga todos los archivos Excel disponibles, evitando duplicados"""
    print("🔍 Buscando archivos Excel recientes automáticamente...")
    
    try:
        selectors = [
            '//div[@data-testid="conversation-panel-body"]',
            '//div[contains(@class, "copyable-area")]//div[contains(@class, "_3YewW")]',
            '//div[contains(@class, "message-list")]',
            '//div[contains(@class, "conversation-panel-messages")]',
            '//div[@role="application"]//div[@role="region"]',
            '//div[@data-testid="chat-list"]',
            '//div[contains(@class, "_1-FMR")]'
        ]
        
        chat_container = None
        for selector in selectors:
            try:
                chat_container = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, selector))
                )
                print(f"✅ Contenedor de mensajes encontrado con selector: {selector}")
                break
            except:
                continue
                
        if not chat_container:
            try:
                script = """
                    return Array.from(document.querySelectorAll('div')).find(el => 
                        el.scrollHeight > 1000 && 
                        el.clientHeight > 300 && 
                        el.scrollHeight > el.clientHeight
                    );
                """
                chat_container = driver.execute_script(script)
                if chat_container:
                    print("✅ Contenedor encontrado con JavaScript alternativo")
                else:
                    raise Exception("No se pudo encontrar el contenedor de mensajes")
            except Exception as js_error:
                print(f"❌ Error con búsqueda JavaScript: {str(js_error)}")
                raise Exception("No se pudo encontrar el contenedor de mensajes")
        
        print("👀 Analizando mensajes recientes en la vista actual...")
        total_downloads = buscar_y_descargar_archivos_en_vista_actual()
        
        scrolls_sin_resultados = 0
        scroll_count = 0
        
        while scrolls_sin_resultados < MAX_SCROLLS_SIN_RESULTADOS and scroll_count < MAX_SCROLLS:
            print(f"🔍 Haciendo scroll para buscar más archivos... (scroll #{scroll_count + 1})")
            
            try:
                driver.execute_script("""
                    arguments[0].scrollTop = arguments[0].scrollHeight * 0.7;
                """, chat_container)
                time.sleep(2.5)
                
                nuevos_downloads = buscar_y_descargar_archivos_en_vista_actual()
                
                total_downloads += nuevos_downloads
                scroll_count += 1
                
                if nuevos_downloads == 0:
                    scrolls_sin_resultados += 1
                    print(f"⚠️ Scroll sin resultados ({scrolls_sin_resultados}/{MAX_SCROLLS_SIN_RESULTADOS})")
                else:
                    scrolls_sin_resultados = 0
                    print(f"✅ Encontrados {nuevos_downloads} nuevos archivos en este scroll")
                
            except Exception as scroll_error:
                print(f"⚠️ Error al hacer scroll: {str(scroll_error)}")
                scrolls_sin_resultados += 1
        
        if scrolls_sin_resultados >= MAX_SCROLLS_SIN_RESULTADOS:
            print(f"✅ Proceso completado: No se encontraron más archivos nuevos después de {scrolls_sin_resultados} intentos")
        elif scroll_count >= MAX_SCROLLS:
            print(f"✅ Proceso completado: Se alcanzó el límite máximo de {MAX_SCROLLS} scrolls")
        
        print(f"✅ Proceso finalizado. Total de archivos descargados: {total_downloads}")
        return total_downloads
    
    except Exception as e:
        print(f"❌ Error durante la búsqueda de archivos: {str(e)}")
        return 0

# =================== EJECUCIÓN PRINCIPAL ===================
try:
    abrir_chat(NOMBRE_CONTACTO)
    
    buscar_descargar_excel_existentes()
    
    print("✅ Bot finalizado correctamente")
    
except Exception as e:
    print(f"❌ Error general: {str(e)}")
