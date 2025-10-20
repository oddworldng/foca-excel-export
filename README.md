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

1. Descarga el ZIP del release.
2. Extrae la carpeta `ExportExcel` completa dentro de `Plugins` de FOCA, quedando:
   `Plugins/ExportExcel/FocaExcelExport.dll`.
3. Reinicia FOCA.

## Uso

1. Abre FOCA.
2. Menú "Plugins" → Load/Unload plugins (carga/descarga de plugins).
   - Verifica que aparece el plugin `ExportExcel`. 
   - Si no aparece, copia la carpeta `ExportExcel` a `Plugins` y continua con el siguiente paso.
3. Menú "Plugins" → "Exportar a Excel".
   - Submenú "Exportar": selecciona proyecto(s) y guarda el Excel.
   - Submenú "Comparar": elige Excel base y Excel nuevo (si el fichero de salida ya existe, se sobrescribe sin preguntar).

## Compilación

1. Abre la solución en Visual Studio (target .NET Framework 4.7.1).
2. Restaura paquetes NuGet (incluye `Fody` y `Costura.Fody`).
3. Compila en Release. La salida es un único `plugin\\ExportExcel\\FocaExcelExport.dll` (dependencias embebidas), listo para comprimir la carpeta `ExportExcel`.

## Estructura del proyecto

```
foca-excel-export/
├── Classes/
│   ├── AssemblyResolver.cs     # Resolución de dependencias (embebidas; fallback a Plugins/lib)
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