

# PST Email Metadata Extractor (Brute Force Mode) 📧

Este script de PowerShell está diseñado para auditar y extraer metadatos de correos electrónicos (Remitente, Destinatarios, Fecha, Message-ID) desde archivos locales de Outlook (`.pst`). 

Su característica principal es el **Modo "Fuerza Bruta"**, diseñado para reconstruir direcciones SMTP válidas incluso cuando Outlook devuelve rutas *Legacy Exchange DN* (ej: `/o=First Organization/ou=.../cn=usuario`) o cuando los metadatos están corruptos.

## 🚀 Características

* **Extracción Recursiva:** Escanea una carpeta raíz y procesa todos los archivos `.pst` encontrados, incluyendo subcarpetas.
* **Recuperación de Direcciones SMTP (Fuerza Bruta):**
    * Intenta obtener la dirección vía propiedades MAPI (`PR_SENDER_SMTP`, `PR_SENT_REPRESENTING`).
    * Si falla, intenta resolver el objeto `ExchangeUser`.
    * Si falla, reconstruye el email parseando la cadena `LegacyExchangeDN` y añadiendo un dominio base detectado del nombre del archivo.
    * Como último recurso, busca patrones de email dentro del nombre visible (Display Name).
* **Deduplicación:** Evita procesar el mismo correo dos veces basándose en el `PR_INTERNET_MESSAGE_ID`.
* **Salida Limpia:** Genera un CSV consolidado con codificación UTF8.

## 📋 Requisitos Previos

* **Sistema Operativo:** Windows 10/11 o Windows Server.
* **Software:** Microsoft Outlook instalado (versión de escritorio, "Classic"). El script utiliza el objeto COM `Outlook.Application`.
* **Permisos:** El usuario que ejecuta el script debe tener permisos de lectura/escritura en las carpetas de los PSTs.

## ⚙️ Configuración

Antes de ejecutar el script, abre el archivo `.ps1` y edita la sección de **CONFIGURACIÓN** al inicio:

```powershell
# --- CONFIGURACIÓN ---
# Ruta donde se encuentran tus archivos .pst
$rutaRaiz = "C:\Ruta\A\Mis\Archivos_PST"

# Ruta donde quieres guardar el reporte final
$archivoSalida = "C:\Ruta\De\Salida\emails_consolidado.csv"
````

También puedes ajustar el **dominio por defecto** en la función `Obtener-DominioDelPST` si el script no logra deducirlo del nombre del archivo:

```powershell
return "tu-empresa.com" # Cambia esto por tu dominio corporativo por defecto
```

## ▶️ Uso

1.  Asegúrate de que **Outlook esté cerrado** (aunque el script intentará instanciarlo, es recomendable no usarlo mientras corre el proceso).
2.  Ejecuta el script desde PowerShell con permisos de administrador (opcional, pero recomendado si accedes a rutas del sistema):

<!-- end list -->

```bash
.\Extract-PSTEmails.ps1
```

3.  El script mostrará el progreso en consola con colores:
      * **Amarillo:** Archivo PST que se está procesando.
      * **Gris:** Progreso de correos (cada 100 emails).
      * **Cian:** Finalización.

## 📊 Salida (CSV)

El archivo generado (`emails_consolidado.csv`) contendrá las siguientes columnas separadas por punto y coma (`;`):

| Columna | Descripción |
| :--- | :--- |
| **MessageID** | Identificador único del correo (Internet Message ID). |
| **From** | Dirección SMTP del remitente (limpia y reconstruida). |
| **To** | Direcciones de los destinatarios separadas por `;`. |
| **DateUTC** | Fecha de envío en formato UTC (`yyyy-MM-dd HH:mm:ss`). |
| **SourcePST** | Nombre del archivo PST de donde se extrajo el dato. |

## ⚠️ Advertencias y Privacidad

  * **Datos Sensibles:** Este script procesa datos confidenciales. Asegúrate de proteger el archivo CSV resultante.
  * **Rendimiento:** El uso de objetos COM de Outlook (MAPI) es intrínsecamente lento comparado con librerías de bajo nivel, pero es más compatible. Para archivos PST de varios gigabytes, el proceso puede tardar horas.
  * **Precisión:** La reconstrucción de emails "Legacy" (Exchange X500) es una aproximación. Si el usuario `cn=juan.perez` ya no existe en la organización o cambió su alias, el email reconstruido `juan.perez@dominio.com` podría no ser funcional, aunque sirve para auditoría histórica.

## 📝 Licencia

Este proyecto está bajo la Licencia [MIT](https://www.google.com/search?q=LICENSE). Siéntete libre de usarlo y modificarlo.

```

***

### Consejos extra para tu repositorio:
1.  **Nombre del archivo:** Guarda tu script con un nombre limpio, por ejemplo: `Get-PstEmailData.ps1`.
2.  **`.gitignore`:** Asegúrate de crear un archivo `.gitignore` y añadir `*.pst` y `*.csv` para evitar subir accidentalmente los correos de tu empresa o el reporte con datos reales a GitHub.
```
