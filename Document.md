# Documentación — Exportar a Obsidian

## 1. Qué hace

**Exportar a Obsidian** es una aplicación de escritorio para Windows que crea
notas Markdown dentro de un vault de Obsidian a partir de:

- Archivos de una carpeta de OneDrive que ya esté sincronizada en el equipo.
- Mensajes y adjuntos de los buzones y carpetas que Outlook clásico muestre en
  el perfil configurado.

La aplicación reúne los dos orígenes en una interfaz gráfica: se indican las
rutas, se eligen los orígenes y se inicia la exportación. El resultado se guarda
en el vault seleccionado; el contenido de los archivos originales y los mensajes
no se envía a un servicio en la nube por esta aplicación.

> **Importante:** hacer una exportación completa puede copiar muchos mensajes,
> archivos y adjuntos. Prueba primero con una carpeta pequeña y un vault de
> prueba. Comprueba el destino antes de iniciar una exportación de correo.

## 2. Requisitos

- Windows con Windows PowerShell 5.1 y .NET Framework.
- Un vault de Obsidian existente, o permisos para crear una carpeta de destino.
- **Para el correo:** Outlook clásico de escritorio instalado, con un perfil
  iniciado y los buzones que quieras exportar disponibles. Outlook nuevo no es
  compatible con la automatización utilizada.
- **Para Word, PDF, Excel o PowerPoint:** la aplicación de Microsoft Office
  correspondiente instalada y disponible en el equipo.
- **Para otros formatos compatibles:** no se necesita Office.

La aplicación **no** inicia sesión en Microsoft Graph, no pide una contraseña
de Microsoft y no descarga contenido de OneDrive. Los archivos que solo existen
en línea deben estar disponibles localmente o descargarse con OneDrive antes de
la exportación.

## 3. Obtener y ejecutar

### Ejecutable incluido en este repositorio

Abre `dist\ExportarAObsidian.exe`. La primera vez, Windows puede mostrar una
advertencia porque el ejecutable no tiene una firma de código.

### Ejecutar directamente el script

Desde PowerShell abierto en la carpeta del proyecto:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\ExportarAObsidian.ps1
```

### Reconstruir el ejecutable

El proyecto incluye el código fuente del lanzador y de la aplicación. Para
volver a generar el ejecutable desde PowerShell:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
```

El resultado aparece en `dist\ExportarAObsidian.exe`. La compilación utiliza
el compilador de C# del .NET Framework de Windows y no requiere instalar
paquetes externos. `-ExecutionPolicy Bypass` se aplica únicamente al proceso
de PowerShell de ese comando; no cambia la política permanente del equipo.

## 4. Uso paso a paso

1. Inicia el ejecutable y deja que se abra la ventana **Exportar a Obsidian**.
2. En **Carpeta local de OneDrive sincronizada**, selecciona la carpeta que
   contiene los documentos que quieres procesar. La aplicación intenta
   sugerir la ubicación de OneDrive del usuario actual; comprueba que sea la
   cuenta y la carpeta correctas.
3. En **Carpeta destino del vault de Obsidian**, elige la raíz del vault
   existente. También se puede elegir una carpeta vacía y añadirla después
   como vault en Obsidian.
4. Marca las casillas correspondientes a los orígenes que quieres exportar.
   Ambas opciones están activadas inicialmente.
5. Si exportas correo, confirma que Outlook clásico está configurado con el
   perfil correcto. Se procesan todos los almacenes y las carpetas
   disponibles; no hay un filtro para elegir solo una carpeta, remitente o
   periodo.
6. Pulsa **Iniciar exportación** y sigue los mensajes de la ventana.
7. Revisa el resumen final y el archivo `ExportarAObsidian.log` en la raíz
   del vault. Si hay errores u omisiones, revísalos antes de dar por
   terminada la exportación.
8. Abre o vuelve a cargar el vault en Obsidian. Las notas nuevas se
   reconocerán como archivos Markdown normales.

La conversión se ejecuta en el hilo de la interfaz. Para buzones o carpetas muy
grandes puede tardar; la ventana muestra el avance y el registro en tiempo real.
No cierres la ventana hasta que aparezca el resumen final.

## 5. Estructura de salida

La aplicación conserva los subdirectorios de OneDrive y separa los dos
orígenes:

```text
<vault>\
|-- OneDrive\
|   |-- <subcarpetas originales>\
|   |   `-- <nombre del archivo>.md
|   `-- <otros archivos convertidos>.md
|-- Outlook\
|   |-- <nombre del almacen>\
|   |   `-- <carpeta de Outlook>\
|   |       |-- <subcarpeta>\
|   |       `-- <año>\
|   |           |-- AAAA-MM-DD_HHMMSS - <asunto>.md
|   |           `-- Adjuntos\
|   |               `-- <nombre de la nota>\
|   |                   `-- <archivo adjunto>
|-- ExportarAObsidian.log
`-- <notas y archivos que ya existían en el vault>
```

La sección `Outlook` es recursiva: las carpetas y subcarpetas se representan
en la estructura del vault. Los nombres de las notas de correo incluyen la
fecha y hora recibidas y el asunto. Los archivos no compatibles no se copian.

### Metadatos de las notas

Las notas de OneDrive incluyen propiedades YAML con `tipo`, `archivo`,
`origen` y `modificado`. Las notas de correo incluyen `tipo`, `asunto`, `de`,
`para` y `fecha`. Los valores de estas propiedades se serializan como cadenas
para conservar caracteres especiales. El cuerpo del mensaje se escribe en la
sección `Mensaje`; si hay adjuntos, se añade una sección `Adjuntos` con enlaces
a los archivos guardados.

## 6. Formatos de OneDrive

| Extensiones | Conversión | Requisitos y límites |
|---|---|---|
| `.txt`, `.md`, `.markdown`, `.json`, `.xml`, `.log` | Se copia el contenido de texto dentro de una nota `.md`. | No requiere Office. |
| `.csv` | Encabezados y filas se representan como una tabla Markdown. | La primera fila debe contener los nombres de las columnas. |
| `.html`, `.htm` | Se eliminan etiquetas HTML y se conservan los saltos y el texto. | Conversión sencilla a texto; no mantiene fielmente estilos ni diseño. |
| `.doc`, `.docx`, `.rtf` | Se extrae el texto del documento. | Requiere Microsoft Word; se pierde parte del formato. |
| `.pdf` | Se extrae el texto que Word pueda abrir. | Requiere Microsoft Word; la calidad depende del PDF y no realiza OCR de documentos escaneados. |
| `.xls`, `.xlsx`, `.xlsm` | Se extraen los valores mostrados de cada hoja en forma de tabla. | Requiere Microsoft Excel; máximo 300 filas y 30 columnas por hoja. |
| `.ppt`, `.pptx` | Se extraen los textos de las diapositivas. | Requiere Microsoft PowerPoint; no conserva imágenes, gráficos ni el diseño. |
| Resto de extensiones | Se omiten y se registran en el log. | No se generan archivos Markdown. |

La conversión de texto de Word, PDF y PowerPoint no equivale a conservar la
apariencia del documento original. Los archivos grandes, complejos, protegidos
por contraseña, corruptos o bloqueados pueden requerir una revisión manual; los
errores de conversión se anotan en el registro.

## 7. Comportamiento del correo y de los adjuntos

- Se recorren todos los almacenes que Outlook expone al perfil actual y todas
  sus carpetas y subcarpetas disponibles.
- Se exportan los elementos que Outlook identifica como mensajes de correo.
  Los elementos de calendario, tareas, contactos y reuniones no se convierten
  en notas con esta versión.
- El mensaje incluye remitente, destinatarios, fecha, asunto y cuerpo en texto
  plano. Los destinatarios CC se añaden cuando Outlook proporciona ese dato.
- Los adjuntos se guardan en la subcarpeta `Adjuntos` de la nota, no se
  convierten automáticamente a Markdown. Cada archivo se enlaza desde la nota.
- El asunto se utiliza en el nombre del archivo y se adapta a las restricciones
  del sistema de archivos. Un asunto vacío se sustituye por `Sin nombre`.
- No se eliminan ni modifican los mensajes originales de Outlook.
- No se añaden archivos adjuntos al sistema de control de versiones; quedan
  únicamente en el destino de exportación que el usuario seleccione.

## 8. Repeticiones, colisiones y modificaciones

La aplicación no sobrescribe un nombre de salida existente. Si la ruta ya está
ocupada, añade un sufijo, por ejemplo `Informe (2).md`. Esta protección también
se aplica a los nombres de adjuntos.

La protección contra sobrescritura **no es una sincronización**: no detecta que
el mensaje o documento ya fue exportado en una ejecución anterior, no actualiza
la nota anterior si cambia el origen y no elimina exportaciones obsoletas.
Volver a exportar puede generar notas duplicadas con sufijos. Haz copia de
seguridad del vault antes de ejecutar una exportación posterior y comprueba el
resultado antes de eliminar duplicados manualmente.

Si el vault está dentro de la carpeta de OneDrive de origen, los archivos
generados dentro del vault se omiten al recorrer el origen para evitar
reconvertir las notas en iteraciones posteriores.

## 9. Privacidad y seguridad

- Las notas de salida contienen información personal: correos, direcciones,
  texto y archivos adjuntos. Elige un vault con los controles de acceso,
  cifrado y copias de seguridad adecuados a esos datos.
- No selecciones por accidente una carpeta compartida, sincronizada o
  publicada si no tienes autorización para colocar allí información sensible.
- Revisa qué buzones están asociados al perfil de Outlook: todos los
  disponibles para la aplicación pueden entrar en la exportación.
- Los archivos de origen se abren para lectura; la aplicación escribe las
  notas, adjuntos y el registro en el destino que se seleccionó.
- La aplicación no incorpora un filtro de remitentes, fechas, carpetas ni
  tipos de adjunto. Usa un vault de prueba para validar el alcance.
- Los adjuntos procedentes de correos pueden contener archivos maliciosos.
  No los abras ni habilites su contenido activo sin verificarlos con las
  medidas de seguridad de tu organización.
- Este repositorio contiene el código fuente y el ejecutable de la
  aplicación; no incluyas notas exportadas, correos, adjuntos ni registros
  generados dentro del repositorio.
- El ejecutable no tiene firma de código. Si Windows presenta una advertencia,
  comprueba que descargaste la versión del repositorio esperado antes de abrirlo.

## 10. Registro y errores

`ExportarAObsidian.log` se crea en la raíz del vault. Incluye la hora de cada
operación, los archivos convertidos, los elementos omitidos, los errores y el
resumen final: archivos, correos, adjuntos, omisiones y errores. El registro
puede incluir rutas locales y asuntos de correo; protégelo igual que el resto
del vault.

Un archivo incompatible incrementa el contador de omitidos; no significa
necesariamente que la operación general haya fallado. Un error de lectura,
conversión o adjunto aparece con `ERROR` y requiere comprobar el mensaje
asociado en el log.

### Solución de problemas

| Síntoma | Qué comprobar |
|---|---|
| La carpeta de OneDrive no se acepta | Verifica que existe, que la cuenta es la correcta y que el contenido está descargado localmente. |
| No aparecen correos o aparece un error de Outlook | Abre Outlook clásico, inicia sesión y confirma el perfil y los buzones visibles; vuelve a iniciar la aplicación. Outlook nuevo no es compatible. |
| Word, Excel o PowerPoint no se abre | Comprueba que la aplicación correspondiente está instalada, activada y se puede iniciar normalmente en el mismo usuario de Windows. |
| Se omite un documento | Comprueba la extensión en la tabla de formatos. Los formatos no incluidos no se convierten en esta versión. |
| El PDF no contiene texto o el resultado está incompleto | Confirma que Word puede abrir el PDF; los PDF escaneados requieren OCR y no se reconocen automáticamente. |
| Hay errores o faltan archivos | Lee `ExportarAObsidian.log`; verifica permisos de lectura en el origen y permisos de escritura y espacio en el vault. |
| Aparecen notas con sufijos `(2)`, `(3)` | Es una protección contra sobrescritura. Una nueva ejecución no actualiza las notas creadas anteriormente. |
| Windows advierte sobre el ejecutable | El archivo no está firmado. Comprueba su origen y, si lo necesitas, reconstrúyelo a partir del código fuente incluido. |

## 11. Validación técnica

La versión incluida se construyó con el script `build.ps1`. Se comprobaron la
sintaxis de PowerShell, la generación del ejecutable, la apertura y el cierre
de su interfaz y una exportación de prueba con archivos TXT y CSV, incluyendo
subcarpetas, nombres ya existentes y formatos no compatibles. Estas pruebas
no sustituyen una validación con los buzones, documentos y configuración real
del usuario.

No se garantiza una conversión perfecta de todos los archivos de Office ni se
ha ejecutado una exportación de prueba sobre buzones personales. Usa primero
contenido de prueba y revisa el resultado antes de procesar información real.

## 12. Componentes del proyecto

| Ruta | Función |
|---|---|
| `ExportarAObsidian.ps1` | Interfaz, lectura de OneDrive y Outlook, conversiones, guardado y registro. |
| `Launcher.cs` | Lanzador de escritorio que incluye el script y lo ejecuta con Windows PowerShell. |
| `build.ps1` | Compila el ejecutable Windows. |
| `dist\ExportarAObsidian.exe` | Ejecutable compilado distribuido con el proyecto. |
| `README.md` | Resumen y acceso rápido a las instrucciones. |
| `Document.md` | Guía funcional, instalación, uso, limitaciones y resolución de problemas. |

## 13. Mejoras posibles

Esta versión no incluye integración con Microsoft Graph, sincronización
incremental, selección de carpetas o periodos de Outlook, filtrado de
adjuntos, OCR, detección de duplicados entre ejecuciones, registro detallado
por archivo ni configuración de plantillas. El modo nube requeriría registro
de una aplicación y permisos de Microsoft Graph que no forman parte de esta
versión.
