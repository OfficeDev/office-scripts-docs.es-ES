---
title: Ejecución de scripts de Office con Power Automate
description: Cómo obtener scripts de Office para Excel en la Web trabajar con un flujo de trabajo de Power Automate.
ms.date: 05/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85c335eeb736ec544eccb2fbdbe819bdbef6848c
ms.sourcegitcommit: aecbd5baf1e2122d836c3eef3b15649e132bc68e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/16/2022
ms.locfileid: "66128233"
---
# <a name="run-office-scripts-with-power-automate"></a>Ejecución de scripts de Office con Power Automate

[Power Automate](https://flow.microsoft.com) permite agregar scripts de Office a un flujo de trabajo más grande y automatizado. Puede usar Power Automate hacer cosas como agregar el contenido de un correo electrónico a la tabla de una hoja de cálculo o crear acciones en las herramientas de administración de proyectos en función de los comentarios del libro.

## <a name="get-started"></a>Introducción

Si no está familiarizado con Power Automate, se recomienda visitar [Comenzar con Power Automate](/power-automate/getting-started). Allí, puede obtener más información sobre todas las posibilidades de automatización disponibles. Los documentos aquí se centran en cómo funcionan los scripts de Office con Power Automate y cómo esto puede ayudar a mejorar la experiencia de Excel.

Para empezar a combinar scripts de Power Automate y Office, siga el tutorial [Inicio del uso de scripts con Power Automate](../tutorials/excel-power-automate-manual.md). Esto le enseñará a crear un flujo que llame a un script simple. Después de completar ese tutorial y [pasar datos a scripts en un tutorial de flujo de Power Automate ejecución automática](../tutorials/excel-power-automate-trigger.md), vuelva aquí para obtener información detallada sobre cómo conectar scripts de Office a flujos de Power Automate.

## <a name="excel-online-business-connector"></a>Conector de Excel Online (Business)

[Los conectores](/connectors/connectors) son los puentes entre Power Automate y las aplicaciones. El [conector Excel Online (Business)](/connectors/excelonlinebusiness) proporciona a los flujos acceso a Excel libros. La acción "Ejecutar script" le permite llamar a cualquier script Office accesible a través del libro seleccionado. También puede proporcionar los parámetros de entrada de los scripts para que el flujo pueda proporcionar datos o que el script devuelva información para los pasos posteriores del flujo.

> [!IMPORTANT]
> La acción "Ejecutar script" proporciona a las personas que usan el conector de Excel acceso significativo al libro y a sus datos. Además, hay riesgos de seguridad con scripts que realizan llamadas API externas, como se explica en [Llamadas externas desde Power Automate](external-calls.md). Si el administrador está preocupado por la exposición de datos altamente confidenciales, puede desactivar el conector de Excel Online o restringir el acceso a Office Scripts a través de [los controles de administrador de scripts de Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

> [!IMPORTANT]
> Power Automate **no** admite scripts almacenados en SharePoint en este momento.

## <a name="data-transfer-in-flows-for-scripts"></a>Transferencia de datos en flujos para scripts

Power Automate permite pasar fragmentos de datos entre los pasos del flujo. Los scripts se pueden configurar para aceptar los tipos de información que necesite y devolver cualquier cosa del libro que desee en el flujo. La entrada para el script se especifica agregando parámetros a la `main` función (además de `workbook: ExcelScript.Workbook`). La salida del script se declara agregando un tipo de valor devuelto a `main`.

> [!NOTE]
> Al crear un bloque "Ejecutar script" en el flujo, se rellenan los parámetros aceptados y los tipos devueltos. Si cambia los parámetros o los tipos devueltos del script, tendrá que volver a hacer el bloque "Ejecutar script" del flujo. Esto garantiza que los datos se analizan correctamente.

En las secciones siguientes se tratan los detalles de la entrada y la salida de los scripts usados en Power Automate. Si quiere un enfoque práctico para aprender este tema, pruebe El paso de [datos a scripts en un tutorial de flujo de Power Automate ejecución automática](../tutorials/excel-power-automate-trigger.md) o explore el escenario de ejemplo [De recordatorios de tareas automatizadas](../resources/scenarios/task-reminders.md).

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parámetros: pasar datos a un script

Toda la entrada de script se especifica como parámetros adicionales para la `main` función. Por ejemplo, si quisiera que un script aceptara un `string` que representa un nombre como entrada, cambiaría la `main` firma a `function main(workbook: ExcelScript.Workbook, name: string)`.

Al configurar un flujo en Power Automate, puede especificar la entrada de script como valores [estáticos, expresiones](/power-automate/use-expressions-in-conditions) o contenido dinámico. Puede encontrar detalles sobre el conector de un servicio individual en la [documentación de Power Automate Connector](/connectors/).

#### <a name="type-restrictions"></a>Restricciones de tipos

Al agregar parámetros de entrada a la función de `main` un script, tenga en cuenta las siguientes restricciones y asignaciones. También se aplican al tipo de valor devuelto del script.

1. El primer parámetro debe ser de tipo `ExcelScript.Workbook`. Su nombre de parámetro no importa.

1. Se admiten los tipos `string`, `number``boolean`, , `unknown`, `object`y `undefined` .

1. Se admiten matrices (tanto `[]` como `Array<T>` estilos) de los tipos enumerados anteriormente. También se admiten matrices anidadas.

1. Los tipos de unión se permiten si son una unión de literales que pertenecen a un único tipo (como `"Left" | "Right"`, no `"Left", 5`). También se admiten uniones de un tipo admitido con undefined (como `string | undefined`).

1. Los tipos de objeto se permiten si contienen propiedades de tipo `string`, , `number``boolean`, matrices admitidas u otros objetos admitidos. En el ejemplo siguiente se muestran objetos anidados que se admiten como tipos de parámetros.

    ```TypeScript
    // The Employee object is supported because Position is also composed of supported types.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

1. Los objetos deben tener su interfaz o definición de clase definida en el script. Un objeto también se puede definir de forma anónima en línea, como en el ejemplo siguiente.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

#### <a name="optional-and-default-parameters"></a>Parámetros opcionales y predeterminados

1. Se permiten parámetros opcionales y se indican con el modificador `?` opcional (por ejemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)`).

1. Se permiten los valores de parámetro predeterminados (por ejemplo `function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`, .

### <a name="return-data-from-a-script"></a>Devolver datos de un script

Los scripts pueden devolver datos del libro que se usarán como contenido dinámico en un flujo de Power Automate. Las [mismas restricciones de tipo enumeradas anteriormente](#type-restrictions) se aplican al tipo de valor devuelto. Para devolver un objeto, agregue la sintaxis del tipo de valor devuelto a la `main` función . Por ejemplo, si quisiera devolver un `string` valor del script, la `main` firma sería `function main(workbook: ExcelScript.Workbook): string`.

## <a name="example"></a>Ejemplo

En la captura de pantalla siguiente se muestra un flujo de Power Automate que se desencadena cada vez que se le asigna un problema [de GitHub](https://github.com/). El flujo ejecuta un script que agrega el problema a una tabla de un libro de Excel. Si hay cinco o más problemas en esa tabla, el flujo envía un recordatorio por correo electrónico.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Editor de flujo de Power Automate que muestra el flujo de ejemplo.":::

La `main` función del script especifica el identificador de problema y el título del problema como parámetros de entrada, y el script devuelve el número de filas de la tabla de problemas.

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

## <a name="see-also"></a>Recursos adicionales

- [Ejecutar scripts mediante un flujo manual de Power Automate](../tutorials/excel-power-automate-manual.md)
- [Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente](../tutorials/excel-power-automate-trigger.md)
- [Devolver datos de un script a un flujo de Power Automate ejecutado automáticamente](../tutorials/excel-power-automate-returns.md)
- [Información de solución de problemas de Power Automate con scripts de Office](../testing/power-automate-troubleshooting.md)
- [Introducción a Power Automate](/power-automate/getting-started)
- [Documentación de referencia del conector de Excel Online (Empresa)](/connectors/excelonlinebusiness/)
