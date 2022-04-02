---
title: Procedimientos recomendados para Scripts de Office
description: Cómo evitar problemas comunes y escribir scripts Office que puedan controlar datos o entradas inesperadas.
ms.date: 12/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 689196e1a0ca70c999ec8048de64190cbfe75581
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585768"
---
# <a name="best-practices-in-office-scripts"></a>Procedimientos recomendados para Scripts de Office

Estos patrones y prácticas están diseñados para ayudar a que los scripts se ejecuten correctamente cada vez. Úsenlos para evitar problemas comunes al empezar a automatizar el flujo Excel trabajo.

## <a name="use-the-action-recorder-to-learn-new-features"></a>Usar la grabadora de acciones para aprender nuevas características

Excel muchas cosas. La mayoría de ellos se pueden crear con scripts. La Grabadora de acciones registra las Excel y las convierte en código. Esta es la forma más sencilla de obtener información sobre cómo funcionan las diferentes características con Office scripts. Si necesita código para una acción específica, cambie a la Grabadora de acciones, realice las acciones, seleccione Copiar como código y pegue el código resultante en el script.

:::image type="content" source="../images/action-recorder-copy-code.png" alt-text="Panel de tareas Grabadora de acciones con el botón &quot;Copiar como código&quot; resaltado.":::

## <a name="verify-an-object-is-present"></a>Comprobar que un objeto está presente

Los scripts suelen basarse en una determinada hoja de cálculo o tabla que está presente en el libro. Sin embargo, pueden cambiar el nombre o quitarse entre las ejecuciones de scripts. Al comprobar si esas tablas o hojas de cálculo existen antes de llamar a métodos en ellas, puede asegurarse de que el script no termine abruptamente.

El siguiente código de ejemplo comprueba si la hoja de cálculo "Índice" está presente en el libro. Si la hoja de cálculo está presente, el script obtiene un rango y procede. Si no está presente, el script registra un mensaje de error personalizado.

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

El operador TypeScript `?` comprueba si el objeto existe antes de llamar a un método. Esto puede simplificar el código si no necesita hacer nada especial cuando el objeto no existe.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Validar primero los datos y el estado del libro

Asegúrese de que todas las hojas de cálculo, tablas, formas y otros objetos estén presentes antes de trabajar en los datos. Con el patrón anterior, compruebe si todo está en el libro y coincide con sus expectativas. Al hacerlo antes de escribir los datos, se asegura de que el script no deje el libro en estado parcial.

El siguiente script requiere que se presenten dos tablas denominadas "Table1" y "Table2". El script comprueba primero si las tablas están presentes y, a continuación, termina con la `return` instrucción y un mensaje adecuado si no lo están.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue...
}
```

Si la comprobación se está produciendo en una función independiente, debe finalizar el script emitiendo la `return` instrucción de la `main` función. La devolución de la subfunción no finaliza el script.

El siguiente script tiene el mismo comportamiento que el anterior. La diferencia es que la función `main` llama a la `inputPresent` función para comprobar todo. `inputPresent` devuelve un valor booleano (`true` o `false`) para indicar si todas las entradas necesarias están presentes. La `main` función usa ese valor booleano para decidir si continúa o finaliza el script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue...
}

function inputPresent(workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a>Cuándo usar una instrucción `throw`

Una [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) instrucción indica que se ha producido un error inesperado. Finaliza el código inmediatamente. En su mayoría, no es necesario desde `throw` el script. Normalmente, el script informa automáticamente al usuario de que el script no se pudo ejecutar debido a un problema. En la mayoría de los casos, basta con finalizar el script con un mensaje de error y una `return` instrucción de la `main` función.

Sin embargo, si el script se ejecuta como parte de un flujo de Power Automate, es posible que desee impedir que el flujo continúe. Una `throw` instrucción detiene el script y le indica al flujo que se detenga también.

El siguiente script muestra cómo usar la `throw` instrucción en nuestro ejemplo de comprobación de tabla.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a>Cuándo usar una instrucción `try...catch`

La [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instrucción es una forma de detectar si se produce un error en una llamada API y seguir ejecutando el script.

Tenga en cuenta el siguiente fragmento de código que realiza una actualización de datos de gran tamaño en un intervalo.

```TypeScript
range.setValues(someLargeValues);
```

Si `someLargeValues` es mayor de Excel para la Web puede controlar, se produce un `setValues()` error en la llamada. A continuación, el script también produce un [error en tiempo de ejecución](../testing/troubleshooting.md#runtime-errors). La `try...catch` instrucción permite que el script reconozca esta condición, sin terminar inmediatamente el script y mostrar el error predeterminado.

Un enfoque para proporcionar al usuario de script una mejor experiencia es presentarles un mensaje de error personalizado. El siguiente fragmento de código muestra una instrucción `try...catch` que registra más información de error para ayudar mejor al lector.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Otro enfoque para tratar los errores es tener un comportamiento de reserva que controle el caso de error. El siguiente fragmento de código usa el `catch` bloque para intentar un método alternativo dividir la actualización en partes más pequeñas y evitar el error.

> [!TIP]
> Para obtener un ejemplo completo sobre cómo actualizar un rango grande, vea [Escribir un conjunto de datos grande](../resources/samples/write-large-dataset.md).

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> Usar `try...catch` dentro o alrededor de un bucle ralentiza el script. Para obtener más información sobre el rendimiento, vea [Evitar el uso de `try...catch` bloques](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).

## <a name="see-also"></a>Consulte también

- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Información de solución de problemas para Power Automate con scripts Office datos](../testing/power-automate-troubleshooting.md)
- [Límites de plataforma con Office scripts](../testing/platform-limits.md)
- [Mejorar el rendimiento de los scripts de Office scripts](web-client-performance.md)
