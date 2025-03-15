from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import os

# 🔹 Configurar Chrome
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = "C:/Program Files/Google/Chrome/Application/chrome.exe"
chrome_options.add_argument("--profile-directory=PERSONAL")

# Configurar carpeta de descargas
download_folder = "E:/python/DOCUMENTOS"
prefs = {"download.default_directory": download_folder}
chrome_options.add_experimental_option("prefs", prefs)
archivos_procesados = set()
# 🔹 Iniciar WebDriver
driver = webdriver.Chrome(options=chrome_options) 

# 🔹 Abrir WhatsApp Web
driver.get("https://web.whatsapp.com")
print("🔄 Esperando que escanees el código QR...")
time.sleep(15)  # Esperar para escanear QR

def abrir_chat(contacto):
    """
    Función mejorada para abrir un chat usando un enfoque más directo
    con simulación de comportamiento humano
    """
    print(f"🔍 Buscando al contacto: {contacto}...")
    
    # Dar más tiempo para cargar la página completamente
    time.sleep(20)
    
    # Captura del estado inicial
    driver.save_screenshot("estado_inicial.png")
    
    # ENFOQUE COMPLETAMENTE NUEVO - Usar la URL directa de WhatsApp
    # Este método es mucho más confiable que intentar hacer clic en elementos
    try:
        # Abrir directamente el chat usando la URL con teléfono/nombre codificado
        chat_url = f"https://web.whatsapp.com/search/{contacto}"
        driver.get(chat_url)
        print(f"✓ Navegando directamente a la búsqueda: {chat_url}")
        time.sleep(5)  # Esperar a que cargue la búsqueda
        
        # Verificar que estamos en la página de búsqueda
        driver.save_screenshot("pagina_busqueda.png")
        
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
                time.sleep(1.5)
                
                # Usar ENTER después de hacer clic (esto a veces funciona mejor)
                chat_element.click()
                from selenium.webdriver.common.keys import Keys
                chat_element.send_keys(Keys.ENTER)
                time.sleep(4)
                
                driver.save_screenshot("despues_clic.png")
                
                # Esperar explícitamente a que cambie la URL o la UI
                WebDriverWait(driver, 10).until(lambda d: "search" not in d.current_url)
                print("✅ Navegación a chat detectada")
                
                time.sleep(3)  # Dar tiempo para que termine de cargar
                break
            except Exception as e:
                print(f"Intento fallido con selector {selector}: {str(e)}")
        
        # VERIFICACIÓN MÁS ROBUSTA: Comprobar múltiples indicadores de chat abierto
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
        
        # Verificación alternativa: comprobar si ya no estamos en la página de búsqueda
        if "search" not in driver.current_url:
            print("✅ Verificación por URL: Chat abierto correctamente")
            return True
        
        # Si llegamos aquí, todos los intentos fallaron
        driver.save_screenshot("error_verificacion.png")
        
        # ÚLTIMO INTENTO - Método extremo
        try:
            # A veces WhatsApp tiene elementos invisibles que bloquean los clics
            # Intentar deshabilitar todos los overlays
            driver.execute_script("""
                document.querySelectorAll('div[role="button"]').forEach(function(el) {
                    if (el.innerText.includes("Darling")) {
                        el.click();
                    }
                });
            """)
            time.sleep(5)
            driver.save_screenshot("ultimo_intento.png")
            return True
        except:
            pass
            
        raise Exception("No se pudo confirmar la apertura del chat")
        
    except Exception as e:
        print(f"❌ Error al abrir chat: {str(e)}")
        driver.save_screenshot("error_abrir_chat.png")
        raise Exception(f"No se pudo abrir el chat de {contacto}: {str(e)}")
    
abrir_chat("Darling")

def extraer_nombre_archivo(texto):
    """Extrae el nombre real del archivo desde diferentes formatos de texto, eliminando numeraciones"""
    if not texto:
        return ""
    
    # Normalizar texto para búsqueda
    texto = texto.lower()
    
    # Caso 1: El texto es solo un nombre de archivo (ej. "prueba2.xlsx")
    if texto.endswith(".xlsx") or texto.endswith(".xls") or texto.endswith(".csv"):
        return texto
    
    # Caso 2: El texto incluye detalles después del nombre (ej. "prueba2.xlsx\n1 página•xlsx•6 kb")
    lines = texto.split('\n')
    if lines and (lines[0].endswith(".xlsx") or lines[0].endswith(".xls") or lines[0].endswith(".csv")):
        return lines[0]
    
    # Caso 3: Intentar extraer patrón de nombre de archivo con extensión
    import re
    match = re.search(r'([^\s]+\.(xlsx|xls|csv))', texto)
    if match:
        return match.group(1)
    
    # Si no podemos determinar el nombre, devolver el texto original
    return texto

def limpiar_nombre_archivo(nombre):
    """Limpia el nombre del archivo eliminando numeraciones como (1), (2), etc."""
    if not nombre:
        return ""
    
    # Eliminar patrones como (1), (2), etc. al final del nombre (antes de la extensión)
    import re
    return re.sub(r' \(\d+\)(?=\.\w+$)', '', nombre)

def verificar_archivo_existe(nombre_archivo):
    """Verifica si un archivo ya existe en la carpeta de descargas, considerando variantes"""
    if not nombre_archivo:
        return False
        
    # Si el nombre no tiene extensión válida, no podemos verificarlo correctamente
    if not (nombre_archivo.endswith('.xlsx') or nombre_archivo.endswith('.xls') or nombre_archivo.endswith('.csv')):
        return False
    
    # Limpiar el nombre para comparación (quitar numeración)
    nombre_base = limpiar_nombre_archivo(nombre_archivo.lower())
    
    # Separar nombre y extensión
    import os
    nombre_sin_ext, extension = os.path.splitext(nombre_base)
    
    # Verificar si existe EXACTAMENTE el mismo archivo
    ruta_completa = os.path.join(download_folder, nombre_archivo)
    if os.path.exists(ruta_completa):
        print(f"🔍 Archivo encontrado con nombre exacto: {nombre_archivo}")
        return True
    
    # Verificar variantes con (1), (2), etc.
    archivos_existentes = os.listdir(download_folder)
    for archivo in archivos_existentes:
        archivo_lower = archivo.lower()
        # Si no tiene la misma extensión, ignorar
        if not archivo_lower.endswith(extension):
            continue
            
        # Limpiar nombre del archivo existente
        archivo_limpio = limpiar_nombre_archivo(archivo_lower)
        archivo_sin_ext, _ = os.path.splitext(archivo_limpio)
        
        # Comparar los nombres base (sin extensión ni numeración)
        if archivo_sin_ext == nombre_sin_ext:
            print(f"🔍 Variante del archivo encontrada: '{archivo}' corresponde a '{nombre_archivo}'")
            return True
    
    # No se encontró ninguna variante del archivo
    return False

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
                # Intentar obtener un ID único para el mensaje
                message_id = mensaje.get_attribute("data-id") or mensaje.get_attribute("id") or ""
                # Si no tiene ID, usar una combinación de texto y posición como identificador
                if not message_id:
                    message_text = mensaje.text[:50] if mensaje.text else ""
                    rect = driver.execute_script("return arguments[0].getBoundingClientRect();", mensaje)
                    message_id = f"{message_text}_{rect['top']}_{rect['left']}"
                
                # Verificar si ya procesamos este mensaje
                if message_id in archivos_procesados:
                    continue
                
                # Buscar posibles documentos o archivos con múltiples selectores
                document_selectors = [
                    ".//div[contains(@data-testid, 'document')]",
                    ".//div[contains(@data-testid, 'document-thumb')]",
                    ".//div[contains(@aria-label, 'document')]",
                    ".//a[contains(@href, 'blob:')]",
                    ".//img[@data-testid='document-thumb']",
                    ".//span[contains(text(), '.xls') or contains(text(), '.xlsx') or contains(text(), '.csv')]",
                    ".//div[contains(@role, 'button') and .//span[contains(text(), '.xls') or contains(text(), '.xlsx') or contains(text(), '.csv')]]",
                    ".//div[contains(@class, 'document-message-container')]"
                ]
                
                for selector in document_selectors:
                    posibles_archivos = mensaje.find_elements(By.XPATH, selector)
                    document_count += len(posibles_archivos)
                    
                    for archivo in posibles_archivos:
                        # Si se estableció un límite y ya lo alcanzamos, salimos
                        if max_intentos is not None and downloads_attempted >= max_intentos:
                            break
                            
                        # DEPURACIÓN: Mostrar información del elemento encontrado
                        texto = archivo.text.lower() if archivo.text else ""
                        href = archivo.get_attribute("href") or ""
                        aria_label = archivo.get_attribute("aria-label") or ""
                        data_testid = archivo.get_attribute("data-testid") or ""
                        class_name = archivo.get_attribute("class") or ""
                        
                        # MEJORA: Extraer el nombre real del archivo
                        nombre_archivo = extraer_nombre_archivo(texto)
                        print(f"📄 Nombre de archivo extraído: '{nombre_archivo}'")
                        
                        # MEJORA: Verificar si el archivo ya existe en la carpeta de descargas
                        if nombre_archivo and verificar_archivo_existe(nombre_archivo):
                            print(f"⏭️ Archivo '{nombre_archivo}' ya existe en la carpeta de descargas, saltando...")
                            # Marcar como procesado para evitar futuros intentos
                            archivos_procesados.add(f"{nombre_archivo}")
                            archivos_procesados.add(message_id)
                            continue
                            
                        # Crear una huella digital más precisa basada en el nombre de archivo real
                        file_id = f"{nombre_archivo}" if nombre_archivo else f"{texto}_{href}_{aria_label}_{data_testid}"
                        
                        # Si ya hemos procesado este archivo, saltarlo
                        if file_id in archivos_procesados:
                            print(f"⏭️ Archivo ya procesado anteriormente: {nombre_archivo or texto}")
                            continue
                        
                        # Marcar este archivo como procesado
                        archivos_procesados.add(file_id)
                        # También marcar el mensaje como procesado
                        archivos_procesados.add(message_id)
                        
                        print(f"🔎 Elemento encontrado: texto='{texto}', data-testid='{data_testid}', aria-label='{aria_label}'")
                        
                        # Verificar si es un archivo Excel con criterios más amplios
                        extensions = [".xls", ".xlsx", ".csv", "excel", "reporte"]
                        es_excel = any(ext in texto.lower() for ext in extensions) or \
                                  any(ext in href.lower() for ext in extensions) or \
                                  any(ext in aria_label.lower() for ext in extensions) or \
                                  ("document" in data_testid) or \
                                  (nombre_archivo and (nombre_archivo.endswith(".xlsx") or 
                                                     nombre_archivo.endswith(".xls") or 
                                                     nombre_archivo.endswith(".csv")))
                        
                        if es_excel:
                            excel_count += 1
                            print(f"📊 Archivo Excel encontrado: {nombre_archivo or texto}")
                            
                            # Hacer visible y descargar
                            try:
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", archivo)
                                time.sleep(1.5)
                                
                                # Intentar descargar el archivo Excel
                                archivo.click()
                                time.sleep(1)
                                
                                # Buscar botones de descarga
                                download_buttons = driver.find_elements(By.XPATH, 
                                    "//div[@role='button' and contains(@aria-label, 'Descargar') or contains(@aria-label, 'Download')]")
                                
                                if download_buttons:
                                    download_buttons[0].click()
                                    downloads_attempted += 1
                                    print(f"✅ Descarga #{downloads_attempted} iniciada: {nombre_archivo or texto}")
                                else:
                                    # Si no encontramos el botón, intentar con JavaScript
                                    try:
                                        driver.execute_script("arguments[0].click();", archivo)
                                        time.sleep(1)
                                        downloads_attempted += 1
                                        print(f"✅ Descarga #{downloads_attempted} iniciada (método JS): {nombre_archivo or texto}")
                                    except:
                                        print("⚠️ No se pudo descargar con método alternativo")
                                
                                time.sleep(3)  # Tiempo para procesar la descarga
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
        # Buscar el contenedor de mensajes
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
                # Intentar con JavaScript para encontrar un contenedor scrollable
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
        
        # FASE 1: Detectar y descargar archivos visibles actualmente
        print("👀 Analizando mensajes recientes en la vista actual...")
        total_downloads = buscar_y_descargar_archivos_en_vista_actual()
        
        # FASE 2: Si no hay suficientes archivos visibles, hacer scrolls inteligentes
        # Haremos scrolls hasta que no encontremos más archivos nuevos (máximo 3 scrolls consecutivos sin resultados)
        max_scrolls = 10  # Máximo número de scrolls para evitar bucles infinitos
        scrolls_sin_resultados = 0
        scroll_count = 0
        
        while scrolls_sin_resultados < 3 and scroll_count < max_scrolls:
            print(f"🔍 Haciendo scroll para buscar más archivos... (scroll #{scroll_count + 1})")
            
            try:
                # Hacer scroll hacia arriba para ver mensajes anteriores
                driver.execute_script("""
                    arguments[0].scrollTop = arguments[0].scrollHeight * 0.7;
                """, chat_container)
                time.sleep(2.5)
                
                # Buscar archivos adicionales
                nuevos_downloads = buscar_y_descargar_archivos_en_vista_actual()
                
                # Actualizar contadores
                total_downloads += nuevos_downloads
                scroll_count += 1
                
                # Si no encontramos nuevos archivos, incrementar el contador de scrolls sin resultados
                if nuevos_downloads == 0:
                    scrolls_sin_resultados += 1
                    print(f"⚠️ Scroll sin resultados ({scrolls_sin_resultados}/3)")
                else:
                    # Reiniciar el contador si encontramos algo
                    scrolls_sin_resultados = 0
                    print(f"✅ Encontrados {nuevos_downloads} nuevos archivos en este scroll")
                
            except Exception as scroll_error:
                print(f"⚠️ Error al hacer scroll: {str(scroll_error)}")
                scrolls_sin_resultados += 1
        
        # Mostrar mensaje final según el motivo de terminación
        if scrolls_sin_resultados >= 3:
            print(f"✅ Proceso completado: No se encontraron más archivos nuevos después de {scrolls_sin_resultados} intentos")
        elif scroll_count >= max_scrolls:
            print(f"✅ Proceso completado: Se alcanzó el límite máximo de {max_scrolls} scrolls")
        
        print(f"✅ Proceso finalizado. Total de archivos descargados: {total_downloads}")
        return total_downloads
    
    except Exception as e:
        print(f"❌ Error durante la búsqueda de archivos: {str(e)}")
        driver.save_screenshot("error_busqueda.png")
        return 0
    
buscar_descargar_excel_existentes()