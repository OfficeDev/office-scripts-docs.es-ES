---
title: Ejecute Office scripts con Power Automate
description: Cómo obtener scripts Office para Excel en la Web trabajar con un flujo Power Automate de trabajo.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: b5bddae61961a56699f99111f71c4f152382f7c6
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/15/2021
ms.locfileid: "59327871"
---
# <a name="run-office-scripts-with-power-automate"></a>Ejecute Office scripts con Power Automate

[Power Automate](https://flow.microsoft.com) permite agregar scripts Office a un flujo de trabajo más grande y automatizado. Puede usar Power Automate tareas como agregar el contenido de un correo electrónico a la tabla de una hoja de cálculo o crear acciones en las herramientas de administración de proyectos en función de los comentarios del libro.

## <a name="get-started"></a>Introducción

Si no es nuevo en Power Automate, le recomendamos que visite [Introducción a Power Automate](/power-automate/getting-started). Allí, puede obtener más información sobre todas las posibilidades de automatización disponibles. Los documentos aquí se centran en cómo Office scripts funcionan con Power Automate y cómo esto puede ayudar a mejorar la experiencia Excel usuario.

Para empezar a combinar Power Automate y Office scripts, siga el tutorial Empezar a usar [scripts con Power Automate](../tutorials/excel-power-automate-manual.md). Esto le enseñará a crear un flujo que llame a un script simple. Después de completar ese tutorial y pasar datos [a scripts](../tutorials/excel-power-automate-trigger.md) en un tutorial de flujo de Power Automate de ejecución automática, vuelva aquí para obtener información detallada acerca de cómo conectar scripts de Office Power Automate flujos.

## <a name="excel-online-business-connector"></a>Excel Conector en línea (empresa)

[Los conectores](/connectors/connectors) son los puentes entre Power Automate y aplicaciones. El [Excel online (empresa)](/connectors/excelonlinebusiness) proporciona a los flujos acceso a Excel libros. La acción "Ejecutar script" le permite llamar a cualquier Office script accesible a través del libro seleccionado. También puede proporcionar parámetros de entrada de scripts para que el flujo pueda proporcionar datos o que el script devuelva información para los pasos posteriores del flujo.

> [!IMPORTANT]
> La acción "Ejecutar script" proporciona a las personas que usan el conector Excel acceso significativo al libro y sus datos. Además, hay riesgos de seguridad con scripts que hacen llamadas API externas, como se explica en [Llamadas externas desde Power Automate](external-calls.md). Si al administrador le preocupa la exposición de datos altamente confidenciales, puede desactivar el conector de Excel Online o restringir el acceso a scripts de Office a través de los controles de administrador de scripts de [Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="data-transfer-in-flows-for-scripts"></a>Transferencia de datos en flujos para scripts

Power Automate permite pasar fragmentos de datos entre los pasos del flujo. Los scripts se pueden configurar para aceptar cualquier tipo de información que necesite y devolver cualquier cosa del libro que desee en el flujo. La entrada del script se especifica agregando parámetros a la `main` función (además de `workbook: ExcelScript.Workbook` ). El resultado del script se declara agregando un tipo devuelto a `main` .

> [!NOTE]
> Al crear un bloque "Ejecutar script" en el flujo, se rellenan los parámetros aceptados y los tipos devueltos. Si cambia los parámetros o devuelve tipos de script, tendrá que volver a hacer el bloque "Ejecutar script" del flujo. Esto garantiza que los datos se analicen correctamente.

Las secciones siguientes cubren los detalles de entrada y salida de scripts usados en Power Automate. Si desea un enfoque práctico para aprender este tema, pruebe el paso de datos [a scripts](../tutorials/excel-power-automate-trigger.md) en un tutorial de flujo de Power Automate de ejecución automática o explore el escenario de ejemplo Avisos de tareas [automatizadas.](../resources/scenarios/task-reminders.md)

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parámetros: pasar datos a un script

Toda la entrada de script se especifica como parámetros adicionales para la `main` función. Por ejemplo, si desea que un script acepte un nombre que represente un nombre como `string` entrada, cambiaría la firma `main` a `function main(workbook: ExcelScript.Workbook, name: string)` .

Al configurar un flujo en Power Automate, puede especificar la entrada de script como valores [estáticos, expresiones](/power-automate/use-expressions-in-conditions)o contenido dinámico. Los detalles sobre el conector de un servicio individual se pueden encontrar en la [documentación Power Automate Connector](/connectors/).

Al agregar parámetros de entrada a la función de un script, tenga en cuenta `main` las siguientes concesiones y restricciones.

1. El primer parámetro debe ser de tipo `ExcelScript.Workbook` . Su nombre de parámetro no importa.

2. Cada parámetro debe tener un tipo (como `string` o `number` ).

3. Los tipos básicos `string` , , , , y se `number` `boolean` `unknown` `object` `undefined` admiten.

4. Se admiten matrices de los tipos básicos enumerados anteriormente.

5. Las matrices anidadas se admiten como parámetros (pero no como tipos devueltos).

6. Los tipos de unión se permiten si son una unión de literales pertenecientes a un único tipo (por `"Left" | "Right"` ejemplo, ). También se admiten uniones de un tipo compatible con undefined (por ejemplo, `string | undefined` ).

7. Los tipos de objeto se permiten si contienen propiedades de `string` tipo , `number` , `boolean` matrices admitidas u otros objetos admitidos. En el ejemplo siguiente se muestran objetos anidados que se admiten como tipos de parámetros:

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. Los objetos deben tener definida su interfaz o definición de clase en el script. Un objeto también se puede definir de forma anónima en línea, como en el ejemplo siguiente:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Los parámetros opcionales se permiten y se pueden anotar como tales mediante el modificador opcional `?` (por ejemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Se permiten los valores de parámetro predeterminados (por `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` ejemplo.

### <a name="return-data-from-a-script"></a>Devolver datos de un script

Los scripts pueden devolver datos del libro que se usarán como contenido dinámico en un flujo Power Automate datos. Al igual que con los parámetros de entrada, Power Automate algunas restricciones en el tipo devuelto.

1. Los tipos básicos `string` , , , y son `number` `boolean` `void` `undefined` compatibles.

2. Los tipos de unión usados como tipos devueltos siguen las mismas restricciones que cuando se usan como parámetros de script.

3. Los tipos de matriz se permiten si son `string` de tipo , o `number` `boolean` . También se permiten si el tipo es una unión admitida o un tipo literal admitido.

4. Los tipos de objeto usados como tipos devueltos siguen las mismas restricciones que cuando se usan como parámetros de script.

5. Se admite la escritura implícita, aunque debe seguir las mismas reglas que un tipo definido.

## <a name="example"></a>Ejemplo

La siguiente captura de pantalla muestra un Power Automate [](https://github.com/) que se desencadena cada vez que se le asigna GitHub un problema. El flujo ejecuta un script que agrega el problema a una tabla de un Excel libro. Si hay cinco o más problemas en esa tabla, el flujo envía un aviso por correo electrónico.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="El Power Automate de flujo que muestra el flujo de ejemplo.":::

La función del script especifica el identificador de problema y el título del problema como parámetros de entrada y el script devuelve el número de `main` filas de la tabla de problemas.

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a>Ver también

- [Ejecute Office scripts en Excel en la Web con Power Automate](../tutorials/excel-power-automate-manual.md)
- [Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente](../tutorials/excel-power-automate-trigger.md)
- [Devolver datos de un script a un flujo de Power Automate ejecutado automáticamente](../tutorials/excel-power-automate-returns.md)
- [Solución de problemas de información Power Automate con scripts Office datos](../testing/power-automate-troubleshooting.md)
- [Introducción a Power Automate](/power-automate/getting-started)
- [Excel Documentación de referencia del conector en línea (empresa)](/connectors/excelonlinebusiness/)
