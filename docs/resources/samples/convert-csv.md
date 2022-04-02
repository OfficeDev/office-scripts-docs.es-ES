---
title: Convertir archivos CSV en Excel libros
description: Obtenga información sobre cómo usar Office scripts y Power Automate para crear .xlsx archivos .csv archivos.
ms.date: 03/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52619c1867b654fae3fce1a383a612f81f80d868
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585593"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>Convertir archivos CSV en Excel libros

Muchos servicios exportan datos como archivos de valores separados por comas (CSV). Esta solución automatiza el proceso de conversión de esos archivos CSV Excel libros en el formato .xlsx archivo. Usa un flujo [Power Automate](https://flow.microsoft.com) para buscar archivos con la extensión .csv en una carpeta de OneDrive y un script de Office para copiar los datos del archivo .csv en un nuevo libro de Excel.

## <a name="solution"></a>Solución

1. Almacene los .csv y un archivo "Template" .xlsx en blanco en una OneDrive carpeta.
1. Cree un script Office para analizar los datos CSV en un intervalo.
1. Cree un flujo Power Automate para leer los .csv y pasar su contenido al script.

## <a name="sample-files"></a>Archivos de ejemplo

Descargue <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> para obtener el archivo Template.xlsx y dos archivos .csv ejemplo. Extraiga los archivos en una carpeta del OneDrive. En este ejemplo se supone que la carpeta se denomina "salida".

Agregue el siguiente script y cree un flujo con los pasos dados para probar el ejemplo usted mismo.

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>Código de ejemplo: insertar valores separados por comas en un libro

```TypeScript
/**
 * Convert incoming CSV data into a range and add it to the workbook.
 */
function main(workbook: ExcelScript.Workbook, csv: string) {
  let sheet = workbook.getWorksheet("Sheet1");

  // Remove any Windows \r characters.
  csv = csv.replace(/\r/g, "");

  // Split each line into a row.
  let rows = csv.split("\n");
  /*
   * For each row, match the comma-separated sections.
   * For more information on how to use regular expressions to parse CSV files,
   * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
   */
  const csvMatchRegex = /(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g
  rows.forEach((value, index) => {
    if (value.length > 0) {
        let row = value.match(csvMatchRegex);
    
        // Check for blanks at the start of the row.
        if (row[0].charAt(0) === ',') {
          row.unshift("");
        }
    
        // Remove the preceding comma.
        row.forEach((cell, index) => {
          row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
        });
    
        // Create a 2D array with one row.
        let data: string[][] = [];
        data.push(row);
    
        // Put the data in the worksheet.
        let range = sheet.getRangeByIndexes(index, 0, 1, data[0].length);
        range.setValues(data);
    }
  });

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate: crear nuevos .xlsx archivos

1. Inicie sesión [Power Automate](https://flow.microsoft.com) y cree un **nuevo flujo de nube programado**.
1. Establece el flujo en **Repetir cada** "1" "Día" y selecciona **Crear**.
1. Obtenga la plantilla Excel archivo. Esta es la base para todos los archivos .csv convertidos. Agregue un **nuevo paso que** use el **conector OneDrive para la Empresa** y la **acción Obtener contenido de** archivo. Proporcione la ruta de acceso al archivo "Template.xlsx".
    * **Archivo**: /output/Template.xlsx
1. Cambie el **nombre del** paso Obtener contenido de archivo yendo al menú Obtener contenido **de archivo (...)** de ese paso (en la esquina superior derecha del conector) y seleccionando la opción **Cambiar** nombre. Cambie el nombre del paso a "Obtener Excel plantilla".

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="El conector OneDrive para la Empresa en Power Automate, cuyo nombre se cambia a Obtener Excel plantilla.":::
1. Obtener todos los archivos de la carpeta "salida". Agregue un **nuevo paso que** use el **conector OneDrive para la Empresa** y los archivos **de lista en la acción de carpeta**. Proporcione la ruta de acceso de carpeta que contiene .csv archivos.
    * **Carpeta**: /output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="El conector OneDrive para la Empresa en Power Automate.":::
1. Agregue una condición para que el flujo solo funcione en .csv archivos. Agregue un **nuevo paso** que sea el control **Condition** . Use los siguientes valores para **condition**.
    * **Elija un valor**: *Nombre* (contenido dinámico de **los archivos de lista de la carpeta**). Tenga en cuenta que este contenido dinámico tiene varios resultados, por lo que un control **Aplicar a cada** control *de* valor rodea la **condición**.
    * **termina con** (de la lista desplegable)
    * **Elija un valor**: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="El control Condition completado con apply to each control around it.":::
1. El resto del flujo se encuentra en la **sección If yes** , ya que solo queremos actuar en .csv archivos. Obtenga un archivo de .csv individual agregando un **nuevo** paso que usa el **conector de** OneDrive para la Empresa y la **acción Obtener contenido de** archivo. Use el **identificador del** contenido dinámico de **los archivos de lista de la carpeta**.
    * **Archivo**: *Id* . (contenido dinámico de los **archivos de lista en el paso de carpeta** )
1. Cambie el nombre **del nuevo paso Obtener contenido de** archivo a "Obtener .csv archivo". Esto ayuda a distinguir este archivo de la Excel plantilla.
1. Realice el nuevo archivo .xlsx, con la plantilla Excel como contenido base. Agregue un **nuevo paso** que use el **conector OneDrive para la Empresa** y la **acción Crear** archivo. Use los siguientes valores.
    * **Ruta de acceso de** carpeta: /output
    * **Nombre de** archivo: *nombre sin* extensión.xlsx (elija el nombre sin  contenido dinámico de extensión de los archivos  de lista en la carpeta y escriba manualmente ".xlsx" después de él)
    * **Contenido de archivo**: *Contenido de archivo* (contenido dinámico de **Obtener Excel plantilla**)

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="Los pasos Obtener .csv archivo y Crear archivo del flujo Power Automate datos.":::
1. Ejecute el script para copiar datos en el nuevo libro. Agregue el **Excel online (empresa)** con la **acción Ejecutar script**. Use los siguientes valores para la acción.
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo**: *Identificador* (contenido dinámico de **Crear archivo**)
    * **Script**: Convertir CSV
    * **csv**: *contenido de archivo* (contenido dinámico de **Obtener .csv archivo**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="El conector Excel online (empresa) completado en Power Automate.":::
1. Guarde el flujo. Use el **botón Probar** en la página del editor de flujo o ejecute el flujo a través de la **pestaña Mis flujos** . Asegúrese de permitir el acceso cuando se le pida.
1. Debe encontrar nuevos archivos .xlsx en la carpeta "salida", junto con los archivos .csv originales. Los libros nuevos contienen los mismos datos que los archivos CSV.

## <a name="troubleshooting"></a>Solución de problemas

### <a name="script-testing"></a>Pruebas de script

Para probar el script sin usar Power Automate, asigne un valor `csv` antes de usarlo. Intente agregar el siguiente código como primera línea de la `main` función y presione **Ejecutar**.

```TypeScript
  csv = `1, 2, 3
         4, 5, 6
         7, 8, 9`;
```

### <a name="semicolon-separated-files-and-other-alternative-separators"></a>Archivos separados por punto y coma y otros separadores alternativos

Algunas regiones usan punto y coma para (';') separar valores de celda en lugar de comas. En este caso, debe cambiar las siguientes líneas en el script.

1. Reemplace las comas por puntos y comas en la instrucción de expresión regular. Esto comienza por `let row = value.match`.

    ```TypeScript
    let row = value.match(/(?:;|\n|^)("(?:(?:"")*[^"]*)*"|[^";\n]*|(?:\n|$))/g);
    ```

1. Reemplace la coma por un punto y coma en la comprobación de la primera celda en blanco. Esto comienza por `if (row[0].charAt(0)`.

    ```TypeScript
    if (row[0].charAt(0) === ';') {
    ```

1. Reemplace la coma por un punto y coma en la línea que quita el carácter de separación del texto mostrado. Esto comienza por `row[index] = cell.indexOf`.

   ```TypeScript
      row[index] = cell.indexOf(";") === 0 ? cell.substr(1) : cell;
    ```

> [!NOTE]
> Si el archivo usa pestañas o cualquier otro carácter para separar los valores, reemplace `;` `\t` las sustituciones anteriores por o cualquier carácter que se esté utilizando.

### <a name="large-csv-files"></a>Archivos CSV grandes

Si el archivo tiene cientos de miles de celdas, podría alcanzar el [Excel de transferencia de datos](../../testing/platform-limits.md#excel). Tendrás que forzar el script para que se sincronice con Excel de forma periódica. La forma más sencilla de hacerlo es llamar después `console.log` de procesar un lote de filas. Agregue las siguientes líneas de código para que esto suceda.

1. Antes `rows.forEach((value, index) => {`, agregue la siguiente línea.

    ```TypeScript
      let rowCount = 0;
    ```

1. Después `range.setValues(data);`de , agregue el siguiente código. Tenga en cuenta que, según el número de columnas, es posible que tenga que reducir `5000` a un número inferior.

    ```TypeScript
      rowCount++;
      if (rowCount % 5000 === 0) {
        console.log("Syncing 5000 rows.");
      }
    ```

> [!WARNING]
> Si el archivo CSV es muy grande, es posible que tenga problemas para realizar [el tiempo de Power Automate](../../testing/platform-limits.md#power-automate). Deberá dividir los datos CSV en varios archivos antes de convertirlos en Excel libros.
