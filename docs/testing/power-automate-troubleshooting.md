---
title: Información de solución de problemas de Power Automate con scripts de Office
description: Sugerencias, información de plataforma y problemas conocidos con la integración entre scripts de Office y Power Automate.
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: 59f4cd8b3476c2ee2a1a862f136173a543ba8a15
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755010"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Información de solución de problemas de Power Automate con scripts de Office

Power Automate le permite llevar la automatización de scripts de Office al siguiente nivel. Sin embargo, dado que Power Automate ejecuta scripts en su nombre en sesiones independientes de Excel, hay algunas cosas importantes que tener en cuenta.

> [!TIP]
> Si acaba de empezar a usar scripts de Office con Power Automate, comience con Ejecutar scripts de Office con [Power Automate](../develop/power-automate-integration.md) para obtener información sobre las plataformas.

## <a name="avoid-using-relative-references"></a>Evitar el uso de referencias relativas

Power Automate ejecuta el script en el libro de Excel elegido en su nombre. Es posible que el libro se cierre cuando esto suceda. Cualquier API que se base en el estado actual del usuario, como `Workbook.getActiveWorksheet` , puede comportarse de forma diferente en Power Automate. Esto se debe a que las API se basan en una posición relativa de la vista o el cursor del usuario y esa referencia no existe en un flujo de Power Automate.

Algunas API de referencia relativas producen errores en Power Automate. Otros tienen un comportamiento predeterminado que implica el estado de un usuario. Al diseñar los scripts, asegúrese de usar referencias absolutas para hojas de cálculo e intervalos. Esto hace que el flujo de Power Automate sea coherente, incluso si las hojas de cálculo se reorganizan.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Métodos de script que fallan al ejecutar flujos de Power Automate

Los siguientes métodos producirán un error y producirán un error cuando se llame desde un script en un flujo de Power Automate.

| Clase | Method |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Métodos de script con un comportamiento predeterminado en flujos de Power Automate

Los siguientes métodos usan un comportamiento predeterminado, en lugar del estado actual de cualquier usuario.

| Clase | Method | Comportamiento de Power Automate |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Devuelve la primera hoja de cálculo del libro o la hoja de cálculo activada actualmente por el `Worksheet.activate` método. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marca la hoja de cálculo como la hoja de cálculo activa para fines de `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Seleccionar libros con el control del explorador de archivos

Al crear el **paso Ejecutar script** de un flujo de Power Automate, debe seleccionar qué libro forma parte del flujo. Use el explorador de archivos para seleccionar el libro, en lugar de escribir manualmente el nombre del libro.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="La acción de script Ejecutar de Power Automate que muestra la opción Mostrar explorador de archivos selector.":::

Para obtener más contexto sobre la limitación de Power Automate y una discusión sobre posibles soluciones alternativas para la selección dinámica de libros, vea este hilo en la comunidad [de Microsoft Power Automate](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Diferencias de zona horaria

Los archivos de Excel no tienen una ubicación o zona horaria inherentes. Cada vez que un usuario abre el libro, su sesión usa la zona horaria local de ese usuario para los cálculos de fecha. Power Automate siempre usa UTC.

Si el script usa fechas u horas, puede haber diferencias de comportamiento cuando el script se prueba localmente frente a cuando se ejecuta a través de Power Automate. Power Automate te permite convertir, dar formato y ajustar tiempos. Consulta [Trabajar](https://flow.microsoft.com/blog/working-with-dates-and-times/) con fechas y horas dentro de los flujos para obtener instrucciones sobre cómo usar esas funciones en Power Automate y [ `main` Parameters: Pasar](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) datos a un script para obtener información sobre cómo proporcionar esa información de hora para el script.

## <a name="see-also"></a>Consulte también

- [Solución de problemas de scripts de Office](troubleshooting.md)
- [Ejecutar scripts de Office con Power Automate](../develop/power-automate-integration.md)
- [Documentación de referencia del conector de Excel Online (Empresa)](/connectors/excelonlinebusiness/)
