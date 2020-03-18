---
title: Leer datos de libros con scripts de Office en Excel en la Web
description: Un tutorial de scripts de Office sobre cómo leer datos de libros y evaluarlos en el script.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700325"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a>Leer datos de libros con scripts de Office en Excel en la Web

Este tutorial le enseñará a leer datos de un libro con un script de Office para Excel en la Web. A continuación, deberá modificar los datos leídos y volver a colocarlos en el libro.

> [!TIP]
> Si no está familiarizado con los scripts de Office, le recomendamos que empiece con el tutorial de [Guardar, editar y crear scripts de Office en Excel en la Web](excel-tutorial.md).

## <a name="prerequisites"></a>Requisitos previos

[!INCLUDE [Preview note](../includes/preview-note.md)]

Antes de iniciar este tutorial, necesitará acceder a los scripts de Office. Esto requiere lo siguiente:

- [Excel en la Web](https://www.office.com/launch/excel)
- Pida a su administrador que [habilite los scripts de Office para su organización](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), lo que agrega la barra de herramientas **Automatizar** a la cinta de opciones.

> [!IMPORTANT]
> Este tutorial está diseñado para las personas con conocimientos a nivel intermedio de JavaScript o TypeScript. Si no está familiarizado con JavaScript, le recomendamos que revise el [Tutorial de JavaScript de Mozilla](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Para obtener más información sobre el entorno de los scripts, visite [Scripts de Office en Excel en la Web](../overview/excel.md).

## <a name="read-a-cell"></a>Leer una celda

Los scripts creados con la Grabadora de acciones solo pueden escribir información en el libro. Con el Editor de código, puede además editar y escribir scripts que lean datos de un libro.

Creemos un script que lea datos y actúe en función de lo que lee. Vamos a trabajar con un ejemplo de extracto bancario. Este ejemplo es una declaración combinada de cuenta corriente y crédito. Lamentablemente, los informes de saldo muestran los cambios de forma diferente. La declaración de cuenta corriente muestra los ingresos como crédito positivo y los costes como débito negativo. En cambio, la declaración de crédito funciona de manera contraria.

En el resto del tutorial, armonizaremos los datos de ambos con un script. En primer lugar, veamos cómo leer los datos del libro.

1. Cree una nueva hoja de cálculo en el libro que ha usado para el resto del tutorial.
2. Copie los siguientes datos y péguelos en la nueva hoja de cálculo, comenzando por la celda **A1**.

    |Fecha |Cuenta |Descripción |Débito |Crédito |
    |:--|:--|:--|:--|:--|
    |10/10/2019 |Cuenta corriente |Coho Vineyard |-20,05 | |
    |11/10/2019 |Crédito |The Phone Company |99,95 | |
    |13/10/2019 |Crédito |Coho Vineyard |154,43 | |
    |15/10/2019 |Cuenta corriente |Depósito externo | |1000 |
    |20/10/2019 |Crédito |Coho Vineyard - Devolución | |- 35,45 |
    |25/10/2019 |Cuenta corriente |Best For You Organics Company | - 85,64 | |
    |01/11/2019 |Cuenta corriente |Depósito externo | |1000 |

3. Abra el **Editor de código** y seleccione **Nuevo script**.
4. Limpiemos un poco el formato. Este es un documento financiero, así que cambie el formato de número de las columnas **Débito** y **Crédito** para mostrar los valores como cantidades en euros. También hay que ajustar el ancho de columna a los datos.

    Reemplace el contenido del script por el siguiente código:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. Ahora, leamos un valor de una de las columnas de número. Agregue el siguiente código al final del script:

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    Observe las llamadas a `load` y `sync`. Puede obtener más información sobre estos métodos en [Aspectos básicos de scripts de Office en Excel en la Web](../develop/scripting-fundamentals.md#sync-and-load). Por ahora, sepa que debe solicitar que se lean los datos y se sincronice el script con el libro para leer esos datos.

6. Ejecute el script.
7. Abra la consola. Vaya al menú de **Puntos suspensivos** y presione **Registros...**.
8. Debe ver `[Array[1]]` en la consola. No es un número por que los rangos son matrices bidimensionales de datos. Este rango bidimensional se ha registrado en la consola directamente. Afortunadamente, el Editor de código le permite ver el contenido de la matriz.
9. Cuando se registra una matriz bidimensional en la consola, se agrupan los valores de columna en cada fila. Expanda el registro de matriz pulsando en el triángulo azul.
10. Expanda el segundo nivel de la matriz pulsando en el triángulo azul que ha descubierto recientemente. Ahora debería ver lo siguiente:

    ![El registro de consola mostrando el resultado "-20,05" anidado en dos matrices.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a>Cambiar el valor de una celda

Ahora que podemos leer datos, vamos a usarlos para modificar el libro. Haremos que el valor de la celda **D2** sea positivo con la función `Math.abs`. El objeto [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contiene varias funciones a las que tienen acceso los scripts. Puede encontrar más información sobre `Math` y otros objetos integrados en [Usar objetos integrados de JavaScript en los scripts de Office](../develop/javascript-objects.md).

1. Agregue el siguiente código al final del script:

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. El valor de la celda **D2** debería ahora ser positivo.

## <a name="modify-the-values-of-a-column"></a>Modificar los valores de una columna

Ahora que sabemos cómo leer y escribir en una sola celda, vamos a aplicar este conocimiento a todas las columnas **Débito** y **Crédito**.

1. Quite el código que afecta a una sola celda (el código de valor absoluto anterior), para que el script tenga el siguiente aspecto:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. Agregue un bucle que itere las filas de las dos últimas columnas. En cada celda, el script establece el valor absoluto del valor actual iterado.

    Tenga en cuenta que la matriz que define las ubicaciones de la celda está basada en cero. Esto significa que la celda **A1** es `range[0][0]`.

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    Esta parte del script realiza varias tareas importantes. En primer lugar, carga los valores y cuenta las filas del rango usado. Esto nos permite ver los valores y averiguar cuándo detenernos. En segundo lugar, itera el rango usado, verificando cada celda en las columnas **Débito** y **Crédito**. Por último, si el valor de la celda no es 0, se reemplaza por su valor absoluto. Evitamos el cero para no cambiar las celdas en blanco.

3. Ejecute el script.

    Ahora, su declaración bancaria debería tener el siguiente aspecto:

    ![Declaración bancaria como tabla con formato que solo contiene valores positivos.](../images/tutorial-5.png)

## <a name="next-steps"></a>Pasos siguientes

Abra el Editor de código y pruebe algunos de nuestros [Ejemplos para scripts de Office en Excel en la Web. Para obtener más información sobre la creación de scripts de Office, consulte también [Fundamentos para scripts de Office en Excel en la Web](../develop/scripting-fundamentals.md).
