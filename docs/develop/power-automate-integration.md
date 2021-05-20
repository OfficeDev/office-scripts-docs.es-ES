---
title: Ejecute scripts de Office con Power Automate
description: Cómo obtener scripts de Office para Excel en la Web trabajando con un flujo de trabajo Power Automate.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7562a2b2359cde67a9a47e0640515018fe23ac35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545043"
---
# <a name="run-office-scripts-with-power-automate"></a>Ejecute scripts de Office con Power Automate

[Power Automate](https://flow.microsoft.com) le permite agregar scripts de Office a un flujo de trabajo automatizado más grande. Puede usar Power Automate hacer cosas como agregar el contenido de un correo electrónico a la tabla de una hoja de trabajo o crear acciones en las herramientas de administración de proyectos en función de los comentarios del libro.

## <a name="get-started"></a>Comenzar

Si eres nuevo en Power Automate, te recomendamos que te [desamiendes Comenzar con Power Automate.](/power-automate/getting-started) Allí, puede obtener más información sobre todas las posibilidades de automatización disponibles para usted. Los documentos aquí se centran en cómo funcionan Office Scripts con Power Automate y cómo eso puede ayudar a mejorar su experiencia Excel.

Para comenzar a combinar scripts de Power Automate y Office, siga el tutorial [Iniciar uso de scripts con Power Automate](../tutorials/excel-power-automate-manual.md). Esto le enseñará cómo crear un flujo que llame a un script simple. Después de completar ese tutorial y pasar datos a scripts en un tutorial [de flujo de Power Automate de ejecución automática,](../tutorials/excel-power-automate-trigger.md) vuelva aquí para obtener información detallada sobre cómo conectar scripts de Office a flujos de Power Automate.

## <a name="excel-online-business-connector"></a>Excel Conector en línea (business)

[Los conectores](/connectors/connectors) son los puentes entre Power Automate y aplicaciones. El [conector Excel online (business)](/connectors/excelonlinebusiness) proporciona a los flujos acceso a Excel libros de trabajo. La acción "Ejecutar script" le permite llamar a cualquier script Office accesible a través del libro seleccionado. También puede proporcionar los parámetros de entrada de scripts para que el flujo pueda proporcionar datos o hacer que el script devuelva información para pasos posteriores en el flujo.

> [!IMPORTANT]
> La acción "Ejecutar script" proporciona a las personas que usan el conector de Excel acceso significativo a su libro de trabajo y sus datos. Además, existen riesgos de seguridad con los scripts que realizan llamadas a la API externa, como se explica en [Llamadas externas desde Power Automate](external-calls.md). Si el administrador está preocupado por la exposición de datos altamente confidenciales, puede desactivar el conector Excel online o restringir el acceso a scripts de Office a través de los [controles de administrador de scripts de Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="data-transfer-in-flows-for-scripts"></a>Transferencia de datos en flujos para scripts

Power Automate le permite pasar fragmentos de datos entre los pasos del flujo. Los scripts se pueden configurar para aceptar los tipos de información que necesite y devolver cualquier cosa de su libro de trabajo que desee en el flujo. La entrada del script se especifica agregando parámetros a la `main` función (además de `workbook: ExcelScript.Workbook` ). La salida del script se declara agregando un tipo de valor devuelto a `main` .

> [!NOTE]
> Al crear un bloque "Ejecutar script" en el flujo, se rellenan los parámetros aceptados y los tipos devueltos. Si cambia los parámetros o los tipos de valor devuelto del script, deberá volver a hacer el bloque "Ejecutar script" del flujo. Esto garantiza que los datos se estén analizando correctamente.

En las secciones siguientes se tratan los detalles de entrada y salida de los scripts utilizados en Power Automate. Si desea un enfoque práctico para aprender este tema, pruebe los [datos de paso a scripts en un](../tutorials/excel-power-automate-trigger.md) tutorial de flujo de Power Automate de ejecución automática o explore el escenario de ejemplo [Recordatorios de tareas automatizadas.](../resources/scenarios/task-reminders.md)

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parámetros: pasar datos a un script

Toda la entrada de script se especifica como parámetros adicionales para la `main` función. Por ejemplo, si desea que un script acepte un `string` que represente un nombre como entrada, cambiaría la `main` firma a `function main(workbook: ExcelScript.Workbook, name: string)` .

Al configurar un flujo en Power Automate, puede especificar la entrada de script como valores [estáticos, expresiones](/power-automate/use-expressions-in-conditions)o contenido dinámico. Los detalles del conector de un servicio individual se pueden encontrar en la documentación del [conector Power Automate.](/connectors/)

Al agregar parámetros de entrada a la función de un `main` script, tenga en cuenta las siguientes asignaciones y restricciones.

1. El primer parámetro debe ser de tipo `ExcelScript.Workbook` . Su nombre de parámetro no importa.

2. Cada parámetro debe tener un tipo (por `string` ejemplo, o `number` ).

3. Los tipos básicos `string` , , , , y se `number` `boolean` `unknown` `object` `undefined` admiten.

4. Se admiten matrices de los tipos básicos enumerados anteriormente.

5. Las matrices anidadas se admiten como parámetros (pero no como tipos de valor devuelto).

6. Los tipos de unión se permiten si son una unión de literales que pertenecen a un único tipo (por `"Left" | "Right"` ejemplo). También se admiten uniones de tipo admitido con indefinidos (por `string | undefined` ejemplo).

7. Los tipos de objeto se permiten si contienen propiedades de tipo `string` `number` , , `boolean` matrices admitidas u otros objetos admitidos. En el ejemplo siguiente se muestran los objetos anidados que se admiten como tipos de parámetro:

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

9. Los parámetros opcionales están permitidos y se pueden denota como tales mediante el modificador opcional `?` (por ejemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Se permiten valores de parámetro predeterminados (por `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` ejemplo.

### <a name="return-data-from-a-script"></a>Devolver datos de un script

Los scripts pueden devolver datos del libro que se usarán como contenido dinámico en un flujo de Power Automate. Al igual que con los parámetros de entrada, Power Automate impone algunas restricciones al tipo de valor devuelto.

1. Los tipos básicos `string` , , , y se `number` `boolean` `void` `undefined` admiten.

2. Los tipos de unión utilizados como tipos de valor devuelto siguen las mismas restricciones que cuando se usan como parámetros de script.

3. Los tipos de matriz se permiten si son de tipo `string` `number` , o `boolean` . También se permiten si el tipo es una unión compatible o un tipo literal admitido.

4. Los tipos de objeto utilizados como tipos de valor devuelto siguen las mismas restricciones que cuando se usan como parámetros de script.

5. Se admite la escritura implícita, aunque debe seguir las mismas reglas que un tipo definido.

## <a name="example"></a>Ejemplo

La siguiente captura de pantalla muestra un flujo de Power Automate que se desencadena cada vez que se le asigna un problema [de GitHub.](https://github.com/) El flujo ejecuta un script que agrega el problema a una tabla de un libro de Excel. Si hay cinco o más problemas en esa tabla, el flujo envía un recordatorio por correo electrónico.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="El editor de flujo Power Automate que muestra el flujo de ejemplo":::

La `main` función del script especifica el id.

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

## <a name="see-also"></a>Vea también

- [Ejecute scripts de Office en Excel en la Web con Power Automate](../tutorials/excel-power-automate-manual.md)
- [Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente](../tutorials/excel-power-automate-trigger.md)
- [Devolver datos de un script a un flujo de Power Automate ejecutado automáticamente](../tutorials/excel-power-automate-returns.md)
- [Información de solución de problemas para Power Automate con scripts de Office](../testing/power-automate-troubleshooting.md)
- [Introducción a Power Automate](/power-automate/getting-started)
- [Excel Documentación de referencia del conector en línea (business)](/connectors/excelonlinebusiness/)
