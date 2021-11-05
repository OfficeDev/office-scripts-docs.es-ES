---
title: Convertir archivos CSV en Excel libros
description: Obtenga información sobre cómo usar Office scripts y Power Automate para crear .xlsx a partir de .csv archivos.
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 203174aec099e426b75d1c816fb3f849b4f13152
ms.sourcegitcommit: 8df930d9ad90001dbed7cb9bd9015ebe7bc9854e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 11/05/2021
ms.locfileid: "60793268"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>Convertir archivos CSV en Excel libros

Muchos servicios exportan datos como archivos de valores separados por comas (CSV). Esta solución automatiza el proceso de conversión de esos archivos CSV Excel libros en el formato .xlsx archivo. Usa un flujo [Power Automate](https://flow.microsoft.com) para buscar archivos con la extensión .csv en una carpeta de OneDrive y un script de Office para copiar los datos del archivo .csv en un nuevo libro de Excel.

## <a name="solution"></a>Solución

1. Almacene los .csv y un archivo "Template" en .xlsx en blanco en una OneDrive carpeta.
1. Cree un Office script para analizar los datos CSV en un intervalo.
1. Cree un flujo Power Automate para leer los archivos .csv y pasar su contenido al script.

## <a name="sample-files"></a>Archivos de ejemplo

Descargue <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> para obtener el archivo Template.xlsx y dos archivos .csv ejemplo. Extraiga los archivos en una carpeta del OneDrive. En este ejemplo se supone que la carpeta se denomina "salida".

Agregue el siguiente script y cree un flujo con los pasos dados para probar el ejemplo usted mismo.

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>Código de ejemplo: insertar valores separados por comas en un libro

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  let data : string[][] = [];
  rows.forEach((value) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);
    
    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });
    data.push(row);
  });

  // Put the data in the worksheet.
  let sheet = workbook.getWorksheet("Sheet1");
  let range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
  range.setValues(data);

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate flujo: crear nuevos .xlsx archivo

1. Inicie sesión [Power Automate](https://flow.microsoft.com) y cree un **nuevo flujo de nube programado.**
1. Establezca el flujo en **Repetir cada** "1" "Día" y seleccione **Crear**.
1. Obtenga la plantilla Excel archivo. Esta es la base para todos los archivos .csv convertidos. Agregue un **nuevo paso** que use el **conector OneDrive para la Empresa** y la acción Obtener **contenido del** archivo. Proporcione la ruta de acceso al archivo "Template.xlsx".
    * **Archivo**: /output/Template.xlsx
1. Cambie el **nombre del** paso Obtener contenido de archivo yendo al menú Obtener contenido de archivo **(...)** de ese paso (en la esquina superior derecha del conector) y seleccionando la opción **Cambiar** nombre. Cambie el nombre del paso a "Obtener Excel plantilla".

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="El conector OneDrive para la Empresa en Power Automate, cuyo nombre se cambia a Obtener Excel plantilla.":::
1. Obtener todos los archivos de la carpeta "salida". Agregue un **nuevo paso que** use el conector **OneDrive para la Empresa** y los archivos de lista en la acción **de carpeta.** Proporcione la ruta de acceso de carpeta que contiene .csv archivos.
    * **Carpeta**: /output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="El conector OneDrive para la Empresa completo en Power Automate.":::
1. Agregue una condición para que el flujo solo funcione en .csv archivos. Agregue un **nuevo paso** que sea el control **Condition.** Use los siguientes valores para **condition**.
    * **Elija un valor**: *Name* (contenido dinámico de archivos de lista en **la carpeta**). Tenga en cuenta que este contenido dinámico tiene varios resultados, por lo que un control Aplicar a **cada** control *de* valor rodea la **condición**.
    * **termina con** (de la lista desplegable)
    * **Elija un valor**: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="El control Condition completado con apply to each control around it.":::
1. El resto del flujo se encuentra en la **sección If yes,** ya que solo queremos actuar en .csv archivos. Obtenga un archivo de .csv individual agregando un **nuevo** paso que usa el **conector de** OneDrive para la Empresa y la acción Obtener **contenido del** archivo. Use el **identificador del** contenido dinámico de los archivos de lista en **la carpeta**.
    * **Archivo:** *Id.* (contenido dinámico del **paso Lista de archivos en carpeta)**
1. Cambie el nombre **del nuevo paso Obtener contenido de** archivo a "Obtener .csv archivo". Esto ayuda a distinguir este archivo de la Excel plantilla.
1. Realice el nuevo archivo .xlsx, usando la plantilla Excel como contenido base. Agregue un **nuevo paso** que use el **conector OneDrive para la Empresa** y la acción **Crear** archivo. Use los siguientes valores.
    * **Ruta de acceso de** carpeta : /output
    * **Nombre de** archivo: *nombre sin* extensión .xlsx (elija el  contenido dinámico Nombre sin extensión de la carpeta Archivos de lista y escriba manualmente ".xlsx" después de él)
    * **Contenido de archivo:** *Contenido de archivo* (contenido dinámico de Obtener Excel **plantilla**)

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="Los pasos Obtener .csv archivo y Crear archivo del flujo Power Automate datos.":::
1. Ejecute el script para copiar datos en el nuevo libro. Agregue el **Excel online (empresa)** con la acción **Ejecutar script.** Use los siguientes valores para la acción.
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo:** *Id.* (contenido dinámico de **Create file**)
    * **Script**: Convertir CSV
    * **csv**: *Contenido de archivo* (contenido dinámico de Obtener .csv **archivo**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="El conector Excel Online (Empresa) completado en Power Automate.":::
1. Guarde el flujo. Use el **botón Probar** en la página del editor de flujo o ejecute el flujo a través de la pestaña **Mis flujos.** Asegúrese de permitir el acceso cuando se le pida.
1. Debe encontrar nuevos archivos .xlsx en la carpeta "salida", junto con los archivos .csv originales. Los libros nuevos contienen los mismos datos que los archivos CSV.

## <a name="troubleshooting"></a>Solución de problemas

El script espera que los valores separados por comas hagan un intervalo rectangular. Si el archivo .csv contiene filas con diferentes números de columnas, aparecerá un error que indica: "El número de filas o columnas de la matriz de entrada no coincide con el tamaño o las dimensiones del rango". Si no se puede hacer que los datos se ajusten a una forma rectangular, use el siguiente script en su lugar. Este script agrega los datos una fila a la vez, en lugar de como un intervalo único. Este script es menos eficaz y es notablemente más lento con conjuntos de datos grandes.

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  let sheet = workbook.getWorksheet("Sheet1");

  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  rows.forEach((value, index) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);

    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });

    // Create a 2D-array with one row.
    let data: string[][] = [];
    data.push(row);

    // Put the data in the worksheet.
    let range = sheet.getRangeByIndexes(index, 0, 1, data[0].length);
    range.setValues(data);
  });

  // Add any formatting or table creation that you want.
}
```
