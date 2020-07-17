---
title: Ejecutar scripts de Office con Power automatization
description: Cómo obtener scripts de Office para Excel en la web trabajar con un flujo de trabajo de Power automatization.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: bd8fea08b7a9303ad2ceace787de6457a33fb979
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160449"
---
# <a name="run-office-scripts-with-power-automate"></a>Ejecutar scripts de Office con Power automatization

La [automatización de energía](https://flow.microsoft.com) permite agregar scripts de Office a un flujo de trabajo más grande y automatizado. Puede usar la función automatizar acciones, como agregar el contenido de un correo electrónico a una tabla de una hoja de cálculo o crear acciones en las herramientas de administración de proyectos en función de los comentarios del libro.

## <a name="getting-started"></a>Introducción

Si es la novedad de la automatización de energía, le recomendamos que visite [Introducción a Power Automate](/power-automate/getting-started). Aquí puede obtener más información sobre todas las posibilidades de automatización disponibles. Los documentos aquí se centran en cómo funcionan los scripts de Office con la automatización energética y cómo esto puede ayudar a mejorar la experiencia de Excel.

Para empezar a combinar la automatización de la alimentación y los scripts de Office, siga el tutorial [comenzar a usar scripts con Power automatization](../tutorials/excel-power-automate-manual.md). Esto le enseñará a crear un flujo que llame a un script sencillo. Una vez que haya completado ese tutorial y los [datos de paso a scripts en un tutorial de flujo automático de energía](../tutorials/excel-power-automate-trigger.md) , vuelva aquí para obtener información detallada sobre la conexión de scripts de Office a la automatización de flujos de alimentación.

## <a name="excel-online-business-connector"></a>Conector de Excel online (Business)

Los [conectores](/connectors/connectors) son los puentes entre las aplicaciones y la automatización de la alimentación. El [conector de Excel online (Business)](/connectors/excelonlinebusiness) proporciona a los flujos acceso a los libros de Excel. La acción "ejecutar script" permite llamar a cualquier script de Office accesible a través del libro seleccionado. No solo puede ejecutar scripts mediante un flujo, sino que puede pasar datos del libro y del flujo de trabajo a través de los scripts.

> [!IMPORTANT]
> La acción "ejecutar script" da a los usuarios que usan el conector de Excel acceso significativo al libro y a sus datos. Además, existen riesgos de seguridad con los scripts que realizan llamadas externas a la API, como se explica en [llamadas externas de la automatización de la alimentación](external-calls.md). Si su administrador está preocupado por la exposición de datos extremadamente confidenciales, puede desactivar el conector de Excel online o restringir el acceso a los scripts de Office a través de los [controles de administrador de scripts de Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).

## <a name="data-transfer-in-flows-for-scripts"></a>Transferencia de datos en flujos para scripts

La automatización de energía permite pasar datos entre los pasos de su flujo. Los scripts se pueden configurar para que acepten los tipos de información que necesite y devuelvan cualquier elemento del libro que desee en su flujo. La entrada para el script se especifica agregando parámetros a la `main` función (además de `workbook: ExcelScript.Workbook` ). El resultado del script se declara agregando un tipo de valor devuelto a `main` .

> [!NOTE]
> Cuando se crea un bloque "Run script" en el flujo, se rellenan los parámetros aceptados y los tipos devueltos. Si cambia los parámetros o tipos de valores devueltos del script, deberá rehacer el bloque "ejecutar script" del flujo. Esto garantiza que los datos se analizan correctamente.

En las secciones siguientes se describen los detalles de entrada y salida de las secuencias de comandos que se usan en la automatización de la energía. Si desea obtener un enfoque práctico para aprender este tema, pruebe [a pasar los datos a los scripts en un tutorial de flujo de automatización de ejecución automática](../tutorials/excel-power-automate-trigger.md) o explorar el escenario de ejemplo de [avisos de tareas automatizadas](../resources/scenarios/task-reminders.md) .

### <a name="main-parameters-passing-data-to-a-script"></a>`main`Parámetros: pasar datos a un script

Todas las entradas de script se especifican como parámetros adicionales para la `main` función. Por ejemplo, si desea que un script acepte un `string` que represente un nombre como entrada, cambiaría la `main` firma a `function main(workbook: ExcelScript.Workbook, name: string)` .

Si está configurando un flujo con la automatización de la alimentación, puede especificar la entrada del script como valores estáticos, [expresiones](/power-automate/use-expressions-in-conditions)o contenido dinámico. Para obtener información sobre el conector de un servicio individual, vaya a la [documentación del conector Power Automated](/connectors/).

Al agregar parámetros de entrada a la función de una secuencia de comandos `main` , tenga en cuenta las siguientes restricciones y concesiones.

1. El primer parámetro debe ser de tipo `ExcelScript.Workbook` . El nombre del parámetro no importa.

2. Cada parámetro debe tener un tipo.

3. `string` `number` `boolean` `any` `unknown` `object` `undefined` Se admiten los tipos básicos,,,,, y.

4. Se admiten las matrices de los tipos básicos enumerados anteriormente.

5. Las matrices anidadas se admiten como parámetros (pero no como tipos devueltos).

6. Los tipos de Unión están permitidos si son una Unión de literales que pertenecen a un tipo único ( `string` , `number` o `boolean` ). También se admiten las uniones de un tipo compatible con undefined.

7. Los tipos de objeto están permitidos si contienen propiedades de tipo `string` ,, `number` `boolean` , matrices admitidas u otros objetos admitidos. En el ejemplo siguiente se muestran los objetos anidados que se admiten como tipos de parámetro:

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

8. Los objetos deben tener su definición de interfaz o clase definida en el script. Un objeto también puede definirse de forma anónima en línea, como en el ejemplo siguiente:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Los parámetros opcionales están permitidos y se pueden marcar como tales mediante el modificador Optional `?` (por ejemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Se permiten los valores predeterminados de parámetro (por ejemplo,) `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .

### <a name="returning-data-from-a-script"></a>Devolución de datos de un script

Los scripts pueden devolver datos del libro que se van a usar como contenido dinámico en un flujo de automatización energética. Al igual que con los parámetros de entrada, la automatización de energía coloca algunas restricciones en el tipo de valor devuelto.

1. Se admiten los tipos básicos `string` , `number` ,, `boolean` `void` y `undefined` .

2. Los tipos de unión usados como tipos de valor devuelto siguen las mismas restricciones que los que se usan cuando se usan como parámetros de script.

3. Los tipos de matriz están permitidos si son del tipo `string` , `number` o `boolean` . También se permiten si el tipo es una Unión compatible o un tipo literal admitido.

4. Los tipos de objeto que se usan como tipos de valor devuelto siguen las mismas restricciones que cuando se usan como parámetros de script.

5. Se admite la escritura implícita, aunque debe seguir las mismas reglas que un tipo definido.

## <a name="avoid-using-relative-references"></a>Evitar el uso de referencias relativas

Power automaticing ejecuta el script en el libro de Excel elegido en su nombre. Es posible que el libro se cierre cuando esto suceda. Cualquier API que se base en el estado actual del usuario, como `Workbook.getActiveWorksheet` , se producirá un error al ejecutarse a través de la automatización de la energía. Al diseñar los scripts, asegúrese de usar referencias absolutas para las hojas de cálculo y los rangos.

Los siguientes métodos producirán un error y no se podrán realizar cuando se llame desde un script en un flujo de automatización de energía.

| Class | Método |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

## <a name="example"></a>Ejemplo

En la siguiente captura de pantalla se muestra un flujo de automatización de energía que se desencadena cuando se le asigna un problema de [GitHub](https://github.com/) . El flujo ejecuta un script que agrega el problema a una tabla de un libro de Excel. Si la tabla tiene cinco o más problemas, el flujo envía un aviso de correo electrónico.

![El flujo de ejemplo, tal como se muestra en el editor de flujo de Power Automate.](../images/power-automate-parameter-return-sample.png)

La `main` función del script especifica el identificador del problema y el título del problema como parámetros de entrada, y el script devuelve el número de filas de la tabla Issue.

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

- [Ejecutar scripts de Office en Excel en la web con la automatización de energía](../tutorials/excel-power-automate-manual.md)
- [Pasar datos a scripts en un flujo de automatización automática de ejecución automática](../tutorials/excel-power-automate-trigger.md)
- [Conceptos básicos de los Scripts de Office en Excel en la web](scripting-fundamentals.md)
- [Introducción a Power Automate](/power-automate/getting-started)
- [Documentación de referencia de Excel online (Business) Connector](/connectors/excelonlinebusiness/)
