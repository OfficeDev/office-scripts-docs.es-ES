---
title: Procedimientos recomendados para Scripts de Office
description: Cómo evitar problemas comunes y escribir scripts de Office robustos que puedan controlar la entrada o los datos inesperados.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546034"
---
# <a name="best-practices-in-office-scripts"></a>Procedimientos recomendados para Scripts de Office

Estos patrones y prácticas están diseñados para ayudar a que los scripts se ejecuten correctamente cada vez. Utilízcalos para evitar trampas comunes a medida que comienza a automatizar el flujo de trabajo de Excel.

## <a name="verify-an-object-is-present"></a>Verificar que un objeto esté presente

Los scripts a menudo se basan en una determinada hoja de cálculo o tabla que está presente en el libro. Sin embargo, es posible que se cambien el nombre o se eliminen entre ejecuciones de script. Al comprobar si esas tablas o hojas de cálculo existen antes de llamar a métodos en ellas, puede asegurarse de que el script no termina abruptamente.

El código de ejemplo siguiente comprueba si la hoja de cálculo "Índice" está presente en el libro. Si la hoja de cálculo está presente, el script obtiene un rango y procede. Si no está presente, el script registra un mensaje de error personalizado.

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

## <a name="validate-data-and-workbook-state-first"></a>Valide primero los datos y el estado del libro de trabajo

Asegúrese de que todas sus hojas de trabajo, tablas, formas y otros objetos estén presentes antes de trabajar en los datos. Usando el patrón anterior, comprueba si todo está en el libro de trabajo y cumple tus expectativas. Hacer esto antes de que se escriban los datos garantiza que el script no deje el libro en un estado parcial.

El siguiente script requiere que dos tablas denominadas "Table1" y "Table2" estén presentes. El script comprueba primero si las tablas están presentes y, a continuación, finaliza con la `return` instrucción y un mensaje adecuado si no lo están.

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

  // Continue....
}
```

Si la verificación se está produciendo en una función independiente, todavía debe finalizar el script emitiendo la `return` instrucción de la `main` función. Volver de la subfunción no termina el script.

El siguiente script tiene el mismo comportamiento que el anterior. La diferencia es que la `main` función llama a la función para verificar `inputPresent` todo. `inputPresent` devuelve un booleano ( `true` o ) para indicar si todas las entradas necesarias están `false` presentes. La `main` función utiliza ese booleano para decidir continuar o finalizar el script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
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

## <a name="when-to-use-a-throw-statement"></a>Cuándo usar una `throw` instrucción

Una [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) instrucción indica que se ha producido un error inesperado. Finaliza el código inmediatamente. En su mayor parte, no es necesario `throw` de su script. Normalmente, el script informa automáticamente al usuario de que el script no se pudo ejecutar debido a un problema. En la mayoría de los casos, es suficiente finalizar el script con un mensaje de error y una `return` instrucción de la `main` función.

Sin embargo, si el script se ejecuta como parte de un flujo de Power Automate, es posible que desee impedir que el flujo continúe. Una `throw` instrucción detiene el script y indica al flujo que también se detenga.

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

## <a name="when-to-use-a-trycatch-statement"></a>Cuándo usar una `try...catch` instrucción

La [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instrucción es una forma de detectar si se produce un error en una llamada a la API y continuar ejecutando el script.

Tenga en cuenta el siguiente fragmento de código que realiza una actualización de datos grande en un intervalo.

```TypeScript
range.setValues(someLargeValues);
```

Si `someLargeValues` es mayor que Excel para la web puede controlar, se produce un error en la `setValues()` llamada. A continuación, el script también falla con un [error de tiempo de ejecución.](../testing/troubleshooting.md#runtime-errors) La `try...catch` instrucción permite que el script reconozca esta condición, sin finalizar inmediatamente el script y mostrar el error predeterminado.

Un enfoque para dar al usuario de script una mejor experiencia es presentarles un mensaje de error personalizado. El siguiente fragmento de código muestra una `try...catch` instrucción que registra más información de error para ayudar mejor al lector.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Otro enfoque para tratar con errores es tener un comportamiento de reserva que controle el caso de error. El siguiente fragmento de código utiliza el `catch` bloque para intentar que un método alternativo rompa la actualización en piezas más pequeñas y evitar el error.

> [!TIP]
> Para obtener un ejemplo completo sobre cómo actualizar un intervalo grande, consulte [Escribir un conjunto de datos grande.](../resources/samples/write-large-dataset.md)

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
> El uso `try...catch` dentro o alrededor de un bucle ralentiza el script. Para obtener más información sobre el rendimiento, consulte [Evitar el uso de `try...catch` bloques.](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)

## <a name="see-also"></a>Vea también

- [Solución de problemas de scripts de Office](../testing/troubleshooting.md)
- [Información de solución de problemas para Power Automate con scripts de Office](../testing/power-automate-troubleshooting.md)
- [Límites de plataforma con scripts de Office](../testing/platform-limits.md)
- [Mejore el rendimiento de sus scripts de Office](web-client-performance.md)
