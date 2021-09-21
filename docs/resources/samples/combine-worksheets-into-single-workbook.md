---
title: Combinar libros en un solo libro
description: Obtenga información sobre cómo usar Office scripts y Power Automate para crear hojas de cálculo de combinación de otros libros en un solo libro.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: ffb0fd13cf587184aec87ade36e5e0e661043b94
ms.sourcegitcommit: c23816babcc628b52f6d8aaa4b6342e04e83a5bd
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/21/2021
ms.locfileid: "59460788"
---
# <a name="combine-worksheets-into-a-single-workbook"></a>Combinar hojas de cálculo en un solo libro

En este ejemplo se muestra cómo extraer datos de varios libros en un único libro centralizado. Usa dos scripts: uno para recuperar información de un libro y otro para crear nuevas hojas de cálculo con esa información. Combina los scripts en un flujo de Power Automate que actúa en una carpeta OneDrive completa.

> [!IMPORTANT]
> En este ejemplo solo se copian los valores de los otros libros. No conserva el formato, los gráficos, las tablas u otros objetos.

## <a name="scenario"></a>Escenario

1. Cree un nuevo archivo Excel en el OneDrive y agregue dos scripts de este ejemplo.
1. Cree una carpeta en el OneDrive y agregue uno o varios libros con datos a él.
1. Cree un flujo para obtener todos los archivos de esa carpeta.
1. Use el script **Devolver datos de hoja de** cálculo para obtener los datos de cada hoja de cálculo en cada uno de los libros.
1. Use el **script Agregar hojas de cálculo** para crear una hoja de cálculo nueva en un solo libro para cada hoja de cálculo de todos los demás archivos.

## <a name="sample-code-return-worksheet-data"></a>Código de ejemplo: Devolver datos de hoja de cálculo

```TypeScript
/**
 * This script returns the values from the used ranges on each worksheet.
 */
function main(workbook: ExcelScript.Workbook): WorksheetData[]
{
  // Create an object to return the data from each worksheet.
  let worksheetInformation: WorksheetData[] = [];

  // Get the data from every worksheet, one at a time.
  workbook.getWorksheets().forEach((sheet) => {
    let values = sheet.getUsedRange()?.getValues();
    worksheetInformation.push({
       name: sheet.getName(),
       data: values as string[][]
    });
  });

  return worksheetInformation;
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="sample-code-add-worksheets"></a>Código de ejemplo: Agregar hojas de cálculo

```TypeScript
/**
 * This script creates a new worksheet in the current workbook for each WorksheetData object provided.
 */
function main(workbook: ExcelScript.Workbook, workbookName: string, worksheetInformation: WorksheetData[])
{
  // Add each new worksheet.
  worksheetInformation.forEach((value) => {
    let sheet = workbook.addWorksheet(`${workbookName}.${value.name}`);

    // If there was any data in the worksheet, add it to a new range.
    if (value.data) {
      let range = sheet.getRangeByIndexes(0, 0, value.data.length, value.data[0].length);
      range.setValues(value.data);
    }
  });
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="power-automate-flow-combine-worksheets-into-a-single-workbook"></a>Power Automate de trabajo: combinar hojas de cálculo en un solo libro

1. Inicie sesión [Power Automate](https://flow.microsoft.com) y cree un nuevo **flujo de nube instantánea.**
1. Elija **Desencadenar manualmente un flujo y** seleccione **Crear**.
1. Obtener todos los archivos de la carpeta. En este ejemplo, usaremos una carpeta denominada "output". Agregue un **nuevo paso que** use el conector **OneDrive para la Empresa** y los archivos de lista en la acción **de carpeta.** Proporcione la ruta de acceso de carpeta que contiene .csv archivos.
    * **Carpeta**: /output

    :::image type="content" source="../../images/combine-worksheets-flow-1.png" alt-text="El conector OneDrive para la Empresa completo en Power Automate.":::
1. Ejecute el script **Devolución de** datos de hoja de cálculo para obtener todos los datos de cada uno de los libros. Agregue el **Excel online (empresa)** con la acción **Ejecutar script.** Use los siguientes valores para la acción. Tenga en cuenta que al agregar el *identificador* del archivo, Power Automate ajustará la acción en un **aplicar** a cada control, por lo que la acción se realizará en cada archivo.
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo:** *Id.* (contenido dinámico de **archivos de lista en la carpeta**)
    * **Script**: Devolver datos de hoja de cálculo
1. Ejecute el **script Agregar hojas de cálculo** en el nuevo Excel archivo que creó. Esto agregará los datos de todos los demás libros. Después de la **acción Ejecutar script** anterior y dentro del control **Aplicar** a cada control, agregue un conector Excel **Online (Empresa)** con la **acción Ejecutar script.** Use los siguientes valores para la acción.
    * **Ubicación**: OneDrive para la Empresa
    * **Biblioteca de documentos**: OneDrive
    * **Archivo:** el archivo
    * **Script**: Agregar hojas de cálculo
    * **workbookName**: *Name* (contenido dinámico de **archivos de lista en la carpeta**)
    * **worksheetInformation** (después de seleccionar el botón **Cambiar** para introducir toda la matriz, vea la nota siguiente a la siguiente imagen): *result* (contenido dinámico del **script Run**)

    :::image type="content" source="../../images/combine-worksheets-flow-2.png" alt-text="Las dos acciones ejecutar script dentro del control Aplicar a cada control.":::
    > [!NOTE]
    > Seleccione el **botón Cambiar para introducir toda la matriz** para agregar el objeto de matriz directamente, en lugar de elementos individuales para la matriz.
    >
    > :::image type="content" source="../../images/combine-worksheets-flow-3.png" alt-text="Botón para cambiar a la entrada de una matriz completa en un cuadro de entrada de campo de control.":::
1. Guarde el flujo. Use el **botón Probar** en la página del editor de flujo o ejecute el flujo a través de la pestaña **Mis flujos.** Asegúrese de permitir el acceso cuando se le pida.
1. El Excel debe tener hojas de cálculo nuevas.
