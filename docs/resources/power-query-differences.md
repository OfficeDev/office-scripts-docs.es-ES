---
title: Cuándo usar Power Query o Scripts de Office
description: Los escenarios más adecuados para las plataformas Power Query y Office scripts.
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: e91077d635d66dde692c129bdd4b2f32657d5283
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585908"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Cuándo usar Power Query o Scripts de Office

[Power Query](https://powerquery.microsoft.com) y Office scripts son soluciones de automatización eficaces para Excel. Ambas soluciones permiten a Excel usuarios limpiar y transformar datos en libros. Un solo script Power Query o Office puede actualizarse y volver a ejecutarse en nuevos datos para producir resultados coherentes, lo que le ahorra tiempo y le permite trabajar con la información resultante más rápido.

En este artículo se proporciona una introducción general sobre cuándo puede favorecer una plataforma sobre la otra. En general, Power Query es bueno para extraer y transformar datos de orígenes de datos externos grandes y scripts de Office son buenos para soluciones [rápidas y centradas](../develop/power-automate-integration.md) en Excel e integraciones Power Automate rápidas.

## <a name="large-data-sources-and-data-retrieval-power-query"></a>Orígenes de datos grandes y recuperación de datos: Power Query

Se recomienda Power Query al tratar con orígenes de datos de plataformas compatibles.

Power Query [conexiones de datos integradas](https://powerquery.microsoft.com/connectors/) a cientos de orígenes. Power Query está especialmente diseñado para tareas de recuperación, transformación y combinación de datos. Cuando necesite datos de uno de esos orígenes, Power Query le ofrece una forma sin código de introducir esos datos en Excel la forma que necesita.

Estas Power Query están diseñadas para conjuntos de datos grandes. No tienen los mismos [límites de transferencia](../testing/platform-limits.md) que Power Automate o Excel para la Web.

Office scripts ofrecen una solución ligera para orígenes de datos o orígenes de datos más pequeños que no están cubiertos por Power Query conectores. Esto incluye [el uso `fetch` de API de REST](../develop/external-calls.md) o la obtención de información de orígenes de datos ad hoc, como una [Teams tarjeta adaptable](../resources/scenarios/task-reminders.md).

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>Formato, visualizaciones y control mediante programación: Office scripts

Se recomienda Office scripts cuando sus necesidades vayan más allá de la importación y transformación de datos.

Casi todo lo que puedes hacer manualmente a través de Excel interfaz de usuario se puede hacer con Office scripts. Son excelentes para aplicar formato coherente a los libros. Los scripts crean gráficos, tablas dinámicas, formas, imágenes y otras visualizaciones de hojas de cálculo. Los scripts también le dan un control preciso sobre las posiciones, tamaños, colores y otros atributos de estas visualizaciones.

La inclusión del código TypeScript proporciona un alto grado de personalización. La lógica de control mediante programación `if...else` , como las instrucciones, hace que el script sea sólido. Esto le permite hacer cosas como leer datos condicionalmente sin depender de fórmulas Excel complejas o examinar el libro en busca de cambios inesperados antes de cambiar el libro.

El formato se puede aplicar con Power Query a través Excel [plantillas](https://templates.office.com/power-query-tutorial-tm11414620). Sin embargo, las plantillas se actualizan en el nivel individual o de la organización, mientras que Office scripts ofrecen un control de acceso más detallado.

## <a name="power-automate-integrations"></a>Power Automate integraciones

Office scripts ofrecen más opciones para la Power Automate integración. Los scripts se adaptan a las soluciones. Defina la entrada [y salida del script](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts), por lo que funciona con cualquier otro conector o datos en el flujo. La siguiente captura de pantalla muestra un ejemplo Power Automate flujo que pasa datos de una Teams adaptable a un script Office usuario.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="Captura de pantalla que muestra el Excel online (empresa) en el diseñador de flujos. El conector usa la acción Ejecutar script para tomar la entrada de una Teams tarjeta adaptable y proporcionarla a un script.":::

Power Query se usa [en el SQL Server](https://powerquery.microsoft.com/flow/) Power Automate conector. La [acción Transformar datos mediante Power Query](/connectors/sql/#transform-data-using-power-query) permite crear una consulta en Power Automate. Aunque se trata de una herramienta eficaz para su uso con SQL Server, limita Power Query a ese origen de entrada, como se muestra en la siguiente captura de pantalla de flujo.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="Captura de pantalla que muestra el SQL Server en el diseñador de flujos. El conector usa los datos Transform mediante Power Query acción.":::

## <a name="platform-dependencies"></a>Dependencias de plataforma

Office scripts solo está disponible actualmente para Excel en la Web. Power Query solo está disponible actualmente para Excel escritorio. Ambos se pueden usar mediante Power Automate, lo que permite que el flujo funcione con Excel libros almacenados en OneDrive.

## <a name="see-also"></a>Consulte también

- [Power Query portal](https://powerquery.microsoft.com/)
- [Power Query con Excel](https://powerquery.microsoft.com/excel/)
- [Ejecutar Office scripts con Power Automate](../develop/power-automate-integration.md)
