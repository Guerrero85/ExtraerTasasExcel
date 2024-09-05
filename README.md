# Procesador de Archivos Excel del Banco Central de Venezuela

Este proyecto de consola en C# lee datos de archivos de Excel, procesa la información, y la guarda en una base de datos SQL Server. A continuación se proporciona una descripción detallada de la funcionalidad del código y las librerías utilizadas.

## Descripción del Código

El propósito principal del código es:

1. **Leer Datos de Archivos Excel**: 
   - El usuario proporciona la ruta de un archivo de Excel.
   - El código lee datos específicos de varias hojas en el archivo de Excel. 
   - Los datos leídos incluyen valores de tasas de cambio para EUR y USD en dos columnas específicas para fechas especificadas en el nombre de la hoja.

2. **Guardar Resultados en un Archivo de Texto**:
   - Los datos extraídos se guardan en un archivo de texto en el directorio de `Downloads` del usuario.

3. **Guardar Datos en una Base de Datos SQL Server**:
   - Los resultados también se insertan o actualizan en una tabla SQL Server.

### Librerías Utilizadas

- **DocumentFormat.OpenXml**: Utilizada para trabajar con archivos de Excel (formatos XLSX). Se necesita instalar el paquete NuGet `DocumentFormat.OpenXml` para manipular documentos de Excel.
- **System.Data.SqlClient**: Utilizada para conectar y operar con una base de datos SQL Server.

### Configuración

1. **Instalación de Dependencias**:
   - Asegúrate de que tu proyecto tenga referencias a las siguientes librerías:
     - `DocumentFormat.OpenXml`
     - `System.Data.SqlClient`

   Puedes instalar `DocumentFormat.OpenXml` usando NuGet Package Manager:

   ```sh
   dotnet add package DocumentFormat.OpenXml



Si deseas contribuir a este proyecto, por favor sigue el proceso estándar de fork y pull request. Asegúrate de probar tus cambios antes de enviarlos.
Este archivo `README.md` proporciona una descripción completa del código y cómo usarlo, configurarlo y ejecutarlo. Puedes adaptarlo según sea necesario para tu proyecto específico.
