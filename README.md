# FocaExcelExport

Plugin de FOCA para exportar a Excel los metadatos de un proyecto: Fichero, URL, Usuario, Ubicación, Email y Cliente. Probado con FOCA Open Source v3.4.7.1.

## Descripción

FocaExcelExport añade al menú de FOCA la opción “Exportar a Excel”. Al seleccionarla se abre un diálogo WinForms que permite elegir un proyecto de la base de datos de FOCA y exportar su información a un fichero `.xlsx` usando ClosedXML (no requiere Excel instalado).

## Funcionalidades

- Exportación a Excel con cabeceras: `Fichero, URL, Usuario, Ubicación, Email, Cliente`.
- Detección dinámica del esquema de la BD (tablas y columnas), compatible con variaciones típicas de FOCA (`FilesITems`, `MetaExtractors`, `EmailsItems/EmailItems`, `UsersItems/UserItems`, etc.).
- Lectura automática del connection string de FOCA vía `ConfigurationManager.ConnectionStrings`.
- Progreso de exportación con barra y mensajes.
- Iconos embebidos en el ensamblado (no se requieren archivos externos para los iconos del menú/botón).

## Requisitos

- FOCA v3.4.7.1
- .NET Framework 4.7.1
- SQL Server (BD de FOCA)

## Instalación

1. Compila el proyecto en Visual Studio (ver “Compilación”).
2. Copia el contenido de la carpeta de salida del proyecto en la carpeta `Plugins` de FOCA. Por defecto el proyecto genera la salida en `plugin\`:
   ```
   Plugins/
     FocaExcelExport.dll
     lib/
       ClosedXML.dll
       DocumentFormat.OpenXml.dll
       ExcelNumberFormat.dll
       System.IO.Packaging.dll
   ```
   Nota: las dependencias se copian automáticamente a `plugin\lib` durante el build y el plugin las carga desde `Plugins/lib` mediante un `AssemblyResolver` propio.
3. Reinicia FOCA.

## Uso

1. Abre FOCA.
2. Menú Plugins → “Exportar a Excel”.
3. Selecciona el proyecto.
4. Pulsa “Exportar” y elige el fichero de destino `.xlsx`.

## Compilación

1. Abre la solución en Visual Studio (target .NET Framework 4.7.1).
2. Restaura paquetes NuGet. Solo se referencia `ClosedXML` de forma directa; el resto se resuelve transitivamente.
3. Compila. La salida se genera en `plugin\` y el target MSBuild copia las DLL necesarias a `plugin\lib`.

## Estructura del proyecto

```
foca-excel-export/
├── Classes/
│   ├── AssemblyResolver.cs     # Carga dependencias desde Plugins/lib
│   ├── ConnectionResolver.cs   # Lee el connection string de FOCA
│   ├── Exporter.cs             # Consulta SQL dinámica y generación del Excel
│   └── SchemaResolver.cs       # Descubre tablas/columnas (INFORMATION_SCHEMA)
├── Forms/
│   ├── ExportDialog.cs         # Diálogo principal (ComboBox, Botón, ProgressBar, Label)
│   └── ExportDialog.Designer.cs
├── Plugin.cs                   # Registro de menú e icono embebido
├── FocaExcelExport.csproj      # Proyecto (.NET Framework 4.7.1)
└── README.md
```