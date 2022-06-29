---
title: Ejecución de scripts de Office con Power Automate
description: Cómo obtener scripts de Office para Excel en la Web trabajar con un flujo de trabajo de Power Automate.
ms.date: 06/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 61e51861bd2c987c25d40e9ac6d2247122256918
ms.sourcegitcommit: c5ffe0a95b962936ee92e7ffe17388bef6d4fad8
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/29/2022
ms.locfileid: "66241859"
---
# <a name="run-office-scripts-with-power-automate"></a>Ejecución de scripts de Office con Power Automate

[Power Automate](https://flow.microsoft.com) le permite agregar scripts de Office a un flujo de trabajo más grande y automatizado. Puede usar Power Automate para hacer cosas como agregar el contenido de un correo electrónico a la tabla de una hoja de cálculo o crear acciones en las herramientas de administración de proyectos en función de los comentarios del libro.

## <a name="get-started"></a>Introducción

Si no está familiarizado con Power Automate, le recomendamos que visite [Introducción a Power Automate](/power-automate/getting-started). Allí, puede obtener más información sobre todas las posibilidades de automatización disponibles. Los documentos aquí se centran en cómo funcionan los scripts de Office con Power Automate y cómo esto puede ayudar a mejorar la experiencia de Excel.

### <a name="step-by-step-tutorials"></a>Tutoriales paso a paso

Hay tres tutoriales paso a paso para Power Automate y Scripts de Office. Estos muestran cómo combinar los servicios de automatización y pasar datos entre un libro y un flujo.

- [Ejecutar scripts mediante un flujo manual de Power Automate](../tutorials/excel-power-automate-manual.md)
- [Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente](../tutorials/excel-power-automate-trigger.md)
- [Devolver datos de un script a un flujo de Power Automate ejecutado automáticamente](../tutorials//excel-power-automate-returns.md)

### <a name="create-a-flow-from-excel"></a>Creación de un flujo desde Excel

Puede empezar a trabajar con Power Automate en Excel con una variedad de plantillas de flujo. En la pestaña **Automatizar** , seleccione **Automatizar una tarea**.

:::image type="content" source="../images/automate-a-task-button.png" alt-text="El botón &quot;Automatizar una tarea&quot; en la cinta de opciones.":::

Se abre un panel de tareas con varias opciones para empezar a conectar los scripts de Office a soluciones automatizadas de mayor tamaño. Seleccione cualquier opción para comenzar. El flujo se proporciona con el libro actual.

:::image type="content" source="../images/automate-a-task-choices.png" alt-text="Panel de tareas que muestra opciones de plantilla de flujo como &quot;Programar un script de Office para ejecutarse en Excel y, a continuación, enviar un correo electrónico&quot; y &quot;Ejecutar un script de Office en Excel cuando se recibe una respuesta Microsoft Forms&quot;.":::

> [!TIP]
> También puede empezar a realizar un flujo desde el menú **Más opciones (...)** de un script individual.

## <a name="excel-online-business-connector"></a>Conector de Excel Online (empresa)

[Los conectores](/connectors/connectors) son los puentes entre Power Automate y las aplicaciones. El [conector de Excel Online (Empresa)](/connectors/excelonlinebusiness) proporciona a los flujos acceso a los libros de Excel. La acción "Ejecutar script" le permite llamar a cualquier script de Office accesible a través del libro seleccionado. También puede proporcionar los parámetros de entrada de los scripts para que el flujo pueda proporcionar datos o que el script devuelva información para los pasos posteriores del flujo.

> [!IMPORTANT]
> La acción "Ejecutar script" proporciona a las personas que usan el conector de Excel acceso significativo al libro y a sus datos. Además, hay riesgos de seguridad con scripts que realizan llamadas API externas, como se explica en [Llamadas externas de Power Automate](external-calls.md). Si el administrador está preocupado por la exposición de datos altamente confidenciales, puede desactivar el conector de Excel Online o restringir el acceso a los scripts de Office a través de [los controles de administrador de Scripts de Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

> [!IMPORTANT]
> Power Automate **no** admite scripts almacenados en SharePoint en este momento.

## <a name="data-transfer-in-flows-for-scripts"></a>Transferencia de datos en flujos para scripts

Power Automate le permite pasar fragmentos de datos entre los pasos del flujo. Los scripts se pueden configurar para aceptar los tipos de información que necesite y devolver cualquier cosa del libro que desee en el flujo. La entrada para el script se especifica agregando parámetros a la `main` función (además de `workbook: ExcelScript.Workbook`). La salida del script se declara agregando un tipo de valor devuelto a `main`.

> [!NOTE]
> Al crear un bloque "Ejecutar script" en el flujo, se rellenan los parámetros aceptados y los tipos devueltos. Si cambia los parámetros o los tipos devueltos del script, tendrá que volver a hacer el bloque "Ejecutar script" del flujo. Esto garantiza que los datos se analizan correctamente.

En las secciones siguientes se tratan los detalles de la entrada y salida de los scripts usados en Power Automate. Si quiere un enfoque práctico para aprender este tema, pruebe el paso de [datos a scripts en un tutorial de flujo de Power Automate de ejecución automática](../tutorials/excel-power-automate-trigger.md) o explore el escenario de ejemplo [De recordatorios de tareas automatizados](../resources/scenarios/task-reminders.md) .

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parámetros: pasar datos a un script

Toda la entrada de script se especifica como parámetros adicionales para la `main` función. Por ejemplo, si quisiera que un script aceptara un `string` que representa un nombre como entrada, cambiaría la `main` firma a `function main(workbook: ExcelScript.Workbook, name: string)`.

Al configurar un flujo en Power Automate, puede especificar la entrada de script como valores [estáticos, expresiones](/power-automate/use-expressions-in-conditions) o contenido dinámico. Puede encontrar detalles sobre el conector de un servicio individual en la [documentación del conector de Power Automate](/connectors/).

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

En la captura de pantalla siguiente se muestra un flujo de Power Automate que se desencadena cada vez que se le asigna un problema de [GitHub](https://github.com/) . El flujo ejecuta un script que agrega el problema a una tabla de un libro de Excel. Si hay cinco o más problemas en esa tabla, el flujo envía un recordatorio por correo electrónico.

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

## <a name="see-also"></a>Vea también

- [Ejecutar scripts mediante un flujo manual de Power Automate](../tutorials/excel-power-automate-manual.md)
- [Pasar datos a scripts en un flujo de Power Automate ejecutado automáticamente](../tutorials/excel-power-automate-trigger.md)
- [Devolver datos de un script a un flujo de Power Automate ejecutado automáticamente](../tutorials/excel-power-automate-returns.md)
- [Información de solución de problemas de Power Automate con scripts de Office](../testing/power-automate-troubleshooting.md)
- [Introducción a Power Automate](/power-automate/getting-started)
- [Documentación de referencia del conector de Excel Online (empresa)](/connectors/excelonlinebusiness/)
