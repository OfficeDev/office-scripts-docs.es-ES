---
title: Solucionar problemas de Office scripts que se ejecutan en Power Automate
description: Sugerencias, la información de la plataforma y los problemas conocidos con la integración entre Office Scripts y Power Automate.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e26378051c764d97b4e8d748abc85fbe095c7b03
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545576"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Solucionar problemas de Office scripts que se ejecutan en Power Automate

Power Automate le permite llevar la automatización de script Office al siguiente nivel. Sin embargo, dado que Power Automate ejecuta scripts en su nombre en sesiones de Excel independientes, hay algunas cosas importantes a tener en cuenta.

> [!TIP]
> Si está empezando a usar scripts de Office con Power Automate, comience con [Ejecutar scripts Office con Power Automate](../develop/power-automate-integration.md) para obtener información sobre las plataformas.

## <a name="avoid-relative-references"></a>Evitar referencias relativas

Power Automate ejecuta el script en el libro de trabajo Excel elegido en su nombre. El libro de trabajo podría cerrarse cuando esto suceda. Cualquier API que se base en el estado actual del usuario, por `Workbook.getActiveWorksheet` ejemplo, puede comportarse de forma diferente en Power Automate. Esto se debe a que las API se basan en una posición relativa de la vista o el cursor del usuario y esa referencia no existe en un flujo Power Automate.

Algunas API de referencia relativas producen errores en Power Automate. Otros tienen un comportamiento predeterminado que implica el estado de un usuario. Al diseñar los scripts, asegúrese de usar referencias absolutas para hojas de trabajo y rangos. Esto hace que su flujo Power Automate sea consistente, incluso si las hojas de trabajo se reorganizan.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Métodos de script que fallan cuando se ejecutan flujos de Power Automate

Los métodos siguientes producirán un error y producirán un error cuando se llama desde un script en un flujo de Power Automate.

| Clase | Método |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Métodos de script con un comportamiento predeterminado en flujos de Power Automate

Los métodos siguientes utilizan un comportamiento predeterminado, en lugar del estado actual de cualquier usuario.

| Clase | Método | comportamiento Power Automate |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Devuelve la primera hoja de cálculo del libro de trabajo o la hoja de cálculo activada actualmente por el `Worksheet.activate` método. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marca la hoja de trabajo como la hoja de trabajo activa para los propósitos de `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Seleccione libros de trabajo con el control del explorador de archivos

Al compilar el paso **Ejecutar script** de un flujo de Power Automate, debe seleccionar qué libro forma parte del flujo. Utilice el explorador de archivos para seleccionar el libro de trabajo, en lugar de escribir manualmente el nombre del libro.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="La acción ejecutar script de Power Automate que muestra la opción Mostrar explorador de archivos selector":::

Para obtener más contexto sobre la limitación de Power Automate y una explicación de las posibles soluciones alternativas para la selección dinámica de libros de trabajo, consulte [este subproceso en microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Diferencias de zona horaria

Excel archivos no tienen una ubicación o zona horaria inherente. Cada vez que un usuario abre el libro, su sesión usa la zona horaria local de ese usuario para los cálculos de fecha. Power Automate siempre usa UTC.

Si el script utiliza fechas o horas, puede haber diferencias de comportamiento cuando el script se prueba localmente en comparación con cuando se ejecuta a través de Power Automate. Power Automate le permite convertir, formatear y ajustar los tiempos. Consulte [Trabajar con fechas y horas dentro de los flujos](https://flow.microsoft.com/blog/working-with-dates-and-times/) para obtener instrucciones sobre cómo usar esas funciones en Power Automate y [ `main` parámetros: pasar datos a un script](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) para obtener información sobre cómo proporcionar esa información de tiempo para el script.

## <a name="see-also"></a>Vea también

- [Solución de problemas de scripts Office](troubleshooting.md)
- [Ejecute scripts de Office con Power Automate](../develop/power-automate-integration.md)
- [Excel Documentación de referencia del conector en línea (business)](/connectors/excelonlinebusiness/)
