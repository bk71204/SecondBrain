# Exportar a Obsidian

Aplicación para Windows que crea notas Markdown en un vault de Obsidian a partir
de archivos de una carpeta sincronizada de OneDrive y correos de Outlook clásico.

## Crear el ejecutable

En Windows, abre PowerShell en esta carpeta y ejecuta:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
```

El ejecutable se genera en `dist\ExportarAObsidian.exe`. También puedes iniciar
la interfaz directamente con:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\ExportarAObsidian.ps1
```

La opción `-ExecutionPolicy Bypass` solo afecta a ese proceso de PowerShell; no
cambia la política del equipo.

## Uso

1. Abre la aplicación y selecciona la carpeta local sincronizada de OneDrive.
2. Selecciona la carpeta del vault de Obsidian que recibirá las notas.
3. Marca los orígenes que quieres exportar e inicia la exportación.
4. Revisa el resumen y `ExportarAObsidian.log` dentro del vault.

Los archivos se organizan bajo `OneDrive\`; los correos, por buzón, carpeta y
año bajo `Outlook\`. Los adjuntos se guardan junto a las notas de correo y se
enlazan desde ellas. Los nombres existentes no se sobrescriben.

## Requisitos y formatos

- Windows PowerShell 5.1, incluido en Windows.
- Outlook clásico de escritorio configurado con un perfil. Outlook nuevo no
  expone la automatización COM que usa esta aplicación.
- Microsoft Word para convertir DOC, DOCX, RTF y PDF; Excel para XLS, XLSX y
  XLSM; PowerPoint para PPT y PPTX. Los archivos TXT, Markdown, CSV, HTML,
  JSON, XML y LOG no requieren Office.
- Se recorren todos los buzones y carpetas disponibles en el perfil de Outlook.
- Los formatos no compatibles se omiten y se registran en el log. Las hojas de
  cálculo se limitan a 300 filas y 30 columnas por hoja.

La aplicación procesa el contenido local de OneDrive, no se conecta a Microsoft
Graph ni descarga por sí misma archivos que OneDrive aún no haya sincronizado.

Para instrucciones completas, requisitos, formatos compatibles, privacidad y
solución de problemas, consulta [Document.md](./Document.md).