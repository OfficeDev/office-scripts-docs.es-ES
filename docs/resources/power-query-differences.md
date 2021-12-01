---
title: Cuándo usar Power Query o Office scripts
description: Los escenarios más adecuados para las plataformas power query y Office scripts.
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1812b508b2cde4d304ecf228adfdd8f68de9808a
ms.sourcegitcommit: 383880e0dc0d09b8f76884675531e462a292d747
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 12/01/2021
ms.locfileid: "61245617"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Cuándo usar Power Query o Office scripts

[Power Query](https://powerquery.microsoft.com) y Office Scripts son soluciones de automatización eficaces para Excel. Ambas soluciones permiten a los Excel limpiar y transformar datos en libros. Un solo script de Power Query o Office puede actualizarse y volver a ejecutarse en nuevos datos para producir resultados coherentes, lo que le ahorra tiempo y le permite trabajar con la información resultante más rápido.

En este artículo se proporciona una introducción general sobre cuándo puede favorecer una plataforma sobre la otra. En general, Power Query es bueno para extraer y transformar datos de grandes orígenes de datos externos y scripts de Office son buenos para soluciones rápidas y [centradas](../develop/power-automate-integration.md)en Excel e integraciones Power Automate.

## <a name="large-data-sources-and-data-retrieval-power-query"></a>Orígenes de datos grandes y recuperación de datos: Power Query

Se recomienda Power Query al tratar con orígenes de datos de plataformas compatibles.

Power Query tiene [conexiones de datos integradas](https://powerquery.microsoft.com/connectors/) a cientos de orígenes. Power Query está especialmente diseñado para tareas de recuperación, transformación y combinación de datos. Cuando necesita datos de uno de esos orígenes, Power Query le ofrece una forma sin código de incorporar esos datos a Excel la forma que necesita.

Estas conexiones de Power Query están diseñadas para conjuntos de datos grandes. No tienen los mismos [límites de transferencia](../testing/platform-limits.md) que Power Automate o Excel para la Web.

Office scripts ofrecen una solución ligera para orígenes de datos más pequeños o orígenes de datos que no están cubiertos por conectores de Power Query. Esto incluye [el uso de API de `fetch` REST](../develop/external-calls.md) o la obtención de información de orígenes de datos ad hoc, como una [Teams tarjeta adaptable.](../resources/scenarios/task-reminders.md)

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>Formato, visualizaciones y control mediante programación: Office scripts

Se recomienda Office scripts cuando sus necesidades vayan más allá de la importación y transformación de datos.

Casi todo lo que puedes hacer manualmente a través de Excel interfaz de usuario se puede hacer con Office scripts. Son excelentes para aplicar formato coherente a los libros. Los scripts crean gráficos, tablas dinámicas, formas, imágenes y otras visualizaciones de hojas de cálculo. Los scripts también le dan un control preciso sobre las posiciones, tamaños, colores y otros atributos de estas visualizaciones.

La inclusión del código TypeScript proporciona un alto grado de personalización. La lógica de control mediante `if...else` programación, como las instrucciones, hace que el script sea sólido. Esto le permite hacer cosas como leer datos condicionalmente sin depender de fórmulas Excel complejas o examinar el libro en busca de cambios inesperados antes de cambiar el libro.

El formato se puede aplicar con Power Query a través Excel [plantillas](https://templates.office.com/power-query-tutorial-tm11414620). Sin embargo, las plantillas se actualizan en el nivel individual o de la organización, mientras que Office scripts ofrecen un control de acceso más detallado.

## <a name="power-automate-integrations"></a>Power Automate integraciones

Office scripts ofrecen más opciones para la Power Automate integración. Los scripts se adaptan a las soluciones. Defina la entrada [y la salida del script,](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts)por lo que funciona con cualquier otro conector o datos en el flujo. En la siguiente captura de pantalla se muestra un ejemplo Power Automate flujo que pasa datos de una Teams tarjeta adaptable a un script Office usuario.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="Captura de pantalla que muestra el Excel online (empresa) en el diseñador de flujos. El conector usa la acción Ejecutar script para tomar la entrada de una Teams tarjeta adaptable y proporcionarla a un script.":::

Power Query se usa en el [SQL Server](https://powerquery.microsoft.com/flow/) Power Automate conector. La [acción Transformar datos mediante Power Query](/connectors/sql/#transform-data-using-power-query) permite crear una consulta en Power Automate. Aunque se trata de una herramienta eficaz para su uso con SQL Server, limita Power Query a ese origen de entrada, como se muestra en la siguiente captura de pantalla de flujo.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="Captura de pantalla que muestra el SQL Server en el diseñador de flujos. El conector usa la acción Transformar datos mediante Power Query.":::

## <a name="platform-dependencies"></a>Dependencias de plataforma

Office Scripts solo está disponible actualmente para Excel en la Web. Power Query actualmente solo está disponible para Excel escritorio. Ambos se pueden usar mediante Power Automate, lo que permite que el flujo funcione con Excel libros almacenados en OneDrive.

## <a name="see-also"></a>Consulte también

- [Power Query Portal](https://powerquery.microsoft.com/)
- [Power Query con Excel](https://powerquery.microsoft.com/excel/)
- [Ejecute Office scripts con Power Automate](../develop/power-automate-integration.md)
