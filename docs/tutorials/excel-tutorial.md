---
title: Grabar, editar y crear scripts de Office en Excel en la Web
description: Un tutorial sobre los conceptos básicos de scripts de Office que incluye la grabación de scripts en la Grabadora de acciones y la escritura de datos en un libro.
ms.date: 05/23/2021
localization_priority: Priority
ms.openlocfilehash: 19cd7bf6c3120d674553d37a36f45d36f46ee852
ms.sourcegitcommit: 0343e4a9843f7ab6ec99d6ddf955050271b061c7
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 05/26/2021
ms.locfileid: "52655908"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a>Grabar, editar y crear scripts de Office en Excel en la Web

En este tutorial se le enseñan los conceptos básicos de la grabación, la edición y la escritura de un script de Office para Excel en la Web. Va a grabar un script que dé formato a una hoja de cálculo con registros de ventas. A continuación, editará el script grabado para dar más formato, crear una tabla y ordenar la tabla. Este método de grabar y luego editar le permite ver el código que resulta de las acciones que ha realizado en Excel.

## <a name="prerequisites"></a>Requisitos previos

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> Este tutorial está diseñado para las personas con conocimientos de nivel intermedio de JavaScript o TypeScript. Si no está familiarizado con JavaScript, le recomendamos que comience con el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Para obtener más información sobre el entorno de los scripts, visite [Entorno del Editor de código de Scripts de Office](../overview/code-editor-environment.md).

## <a name="add-data-and-record-a-basic-script"></a>Agregar datos y grabar un script básico

En primer lugar, necesitaremos algunos datos y un pequeño script inicial.

1. Cree un libro nuevo en Excel para la Web.
2. Copie los siguientes datos de ventas de frutas y péguelos en la hoja de cálculo, comenzando por la celda **A1**.

    |Fruta |2018 |2019 |
    |:---|:---|:---|
    |Naranjas |1000 |1200 |
    |Limones |800 |900 |
    |Limas |600 |500 |
    |Pomelos |900 |700 |

3. Abra la pestaña **Automatizar**. Si no ve la pestaña **Automatizar**, presione la flecha desplegable para comprobar el desbordamiento de la cinta de opciones. Si aún no aparece, siga los consejos del artículo [Solución de problemas de Scripts de Office](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).
4. Haga clic en el botón **Guardar acciones**.
5. Seleccione las celdas **A2:C2** (la fila "Naranjas") y configure el color de relleno como naranja.
6. Detenga la grabación pulsando el botón **Detener**.

    La hoja de cálculo debe tener este aspecto (no se preocupe si el color es diferente):

    :::image type="content" source="../images/tutorial-1.png" alt-text="Una hoja de cálculo donde se muestra la fila de datos de ventas de fruta con la fila que contiene &quot;Naranjas&quot; resaltada en color naranja.":::

## <a name="edit-an-existing-script"></a>Editar un script existente

El script anterior pinta la fila "Naranja" de color naranja. Ahora, agreguemos una fila amarilla a "Limones".

1. En el panel **Detalles** ya abierto, presione el botón **Editar**.
2. Debería ver algo parecido a este código:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    Este código obtiene la hoja de cálculo actual del libro. Después, establece el color de relleno del rango **A2:C2**.

    Los rangos son una parte fundamental de las secuencias de comandos de Office en Excel en la Web. Un rango es un bloque de celdas contiguo y rectangular que contiene valores, fórmulas y formatos. Constituyen la estructura básica de las celdas y se usan para realizar la mayoría de las tareas de scripts.

3. Agregue la línea siguiente al final del script (entre el lugar en el que se establece el `color` y aparece el `}` de cierre):

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. Para probar el script, presione **Ejecutar**. El libro tendrá ahora el siguiente aspecto:

    :::image type="content" source="../images/tutorial-2.png" alt-text=" Una hoja de cálculo donde se muestra una fila de datos de ventas con la fila &quot;Naranjas&quot; resaltada en naranja y la fila &quot;Limones&quot; resaltada en amarillo.":::

## <a name="create-a-table"></a>Crear una tabla

Vamos a convertir estos datos de ventas de frutas en una tabla. Usaremos nuestro script para todo este proceso.

1. Agregue la línea siguiente al final del script (antes del `}` de cierre):

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. Esa llamada devuelve un objeto de `Table`. Vamos a usar la tabla para ordenar los datos. Ordenaremos los datos de menor a mayor en función de los valores de la columna "Frutas". Agregue la siguiente línea después de la creación de tabla:

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    Su script debe tener este aspecto:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    Las tablas tienen un objeto `TableSort` al que se accede mediante el método `Table.getSort`. Puede aplicar un criterio de ordenación a ese objeto. El método `apply` acepta una matriz de objetos `SortField`. En este caso, solo tenemos un criterio de ordenación, por lo que solo usamos un `SortField`. `key: 0` establece los valores que definen la ordenación de la columna como "0" (que es la primera columna de la tabla **A** en este caso). `ascending: true` ordena los datos de menor a mayor (en lugar de mayor a menor).

3. Ejecute el script. Debería ver una tabla como esta:

    :::image type="content" source="../images/tutorial-3.png" alt-text="Una hoja de cálculo donde se muestra la tabla de ventas de frutas ordenadas.":::

    > [!NOTE]
    > Si vuelve a ejecutar el script, se producirá un error. Esto se debe a que no se puede crear una tabla encima de otra. Sin embargo, puede ejecutar el script en otra hoja de cálculo o en un libro.

### <a name="re-run-the-script"></a>Ejecute el script de nuevo

1. Crear una nueva hoja de cálculo en el libro actual.
2. Copie los datos de frutas del principio del tutorial y péguelos en la nueva hoja de cálculo, comenzando en la celda **A1**.
3. Ejecute el script.

## <a name="next-steps"></a>Pasos siguientes

Complete el tutorial [Leer datos de libros con scripts de Office en Excel en la Web](excel-read-tutorial.md). En él aprenderá a leer datos de un libro con un script de Office.
