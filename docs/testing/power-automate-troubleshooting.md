---
title: Solucionar Office scripts que se ejecutan en Power Automate
description: Sugerencias, información de plataforma y problemas conocidos con la integración entre Office scripts y Power Automate.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 028c34003a6f6b00c9afc67450b249b938d445fb
ms.sourcegitcommit: 634ad2061e683ae1032c1e0b55b00ac577adc34f
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 11/03/2021
ms.locfileid: "60725626"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Solucionar Office scripts que se ejecutan en Power Automate

Power Automate permite llevar la automatización Office script al siguiente nivel. Sin embargo, como Power Automate scripts en su nombre en sesiones Excel independientes, hay algunas cosas importantes que tener en cuenta.

> [!TIP]
> Si está empezando a usar scripts de Office con Power Automate, comience con Ejecutar scripts de [Office](../develop/power-automate-integration.md) con Power Automate para obtener información sobre las plataformas.

## <a name="avoid-relative-references"></a>Evitar referencias relativas

Power Automate ejecuta el script en el libro Excel en su nombre. Es posible que el libro se cierre cuando esto suceda. Cualquier API que se base en el estado actual del usuario, como , puede comportarse de forma diferente `Workbook.getActiveWorksheet` en Power Automate. Esto se debe a que las API se basan en una posición relativa de la vista o el cursor del usuario y esa referencia no existe en un flujo Power Automate usuario.

Algunas API de referencia relativas producen errores en Power Automate. Otros tienen un comportamiento predeterminado que implica el estado de un usuario. Al diseñar los scripts, asegúrese de usar referencias absolutas para hojas de cálculo e intervalos. Esto hace que Power Automate flujo de trabajo coherente, incluso si las hojas de cálculo se reorganizan.

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a>Métodos de script que fallan al ejecutarse en Power Automate flujos

Los siguientes métodos inician un error y se produce un error cuando se llama desde un script en un flujo Power Automate datos.

| Clase | Método |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Métodos de script con un comportamiento predeterminado en Power Automate flujos

Los siguientes métodos usan un comportamiento predeterminado, en lugar del estado actual de cualquier usuario.

| Clase | Método | Power Automate comportamiento |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Devuelve la primera hoja de cálculo del libro o la hoja de cálculo activada actualmente por el `Worksheet.activate` método. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marca la hoja de cálculo como la hoja de cálculo activa para fines de `Workbook.getActiveWorksheet` . |

## <a name="data-refresh-not-supported-in-power-automate"></a>No se admite la actualización de datos en Power Automate

Office Los scripts no pueden actualizar los datos cuando se ejecutan en Power Automate. Métodos como `PivotTable.refresh` no hacer nada cuando se llama en un flujo. Además, Power Automate activa una actualización de datos para fórmulas que usan vínculos de libro.

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a>Métodos de script que no hacen nada cuando se ejecutan en Power Automate flujos

Los siguientes métodos no hacen nada en un script cuando se llama a través de Power Automate. Todavía devuelven correctamente y no producen errores.

| Clase | Método |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a>Seleccionar libros con el control del explorador de archivos

Al compilar el **paso Ejecutar script** de un flujo Power Automate, debe seleccionar qué libro forma parte del flujo. Use el explorador de archivos para seleccionar el libro, en lugar de escribir manualmente el nombre del libro.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="La Power Automate ejecutar script que muestra la opción Mostrar explorador de archivos selector.":::

Para obtener más contexto sobre la limitación Power Automate y una discusión sobre posibles soluciones alternativas para la selección dinámica de libros, vea este subproceso en microsoft [Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="pass-entire-arrays-as-script-parameters"></a>Pasar matrices enteras como parámetros de script

Power Automate permite a los usuarios pasar matrices a los conectores como una variable o como elementos individuales de la matriz. El valor predeterminado es pasar elementos individuales, que compilan la matriz en el flujo. Para scripts u otros conectores que toman matrices completas como argumentos, debe seleccionar el botón **Cambiar** para introducir toda la matriz para pasar la matriz como un objeto completo. Este botón está en la esquina superior derecha de cada campo de entrada de parámetro de matriz.

:::image type="content" source="../images/combine-worksheets-flow-3.png" alt-text="Botón para cambiar a la entrada de una matriz completa en un cuadro de entrada de campo de control.":::

## <a name="time-zone-differences"></a>Diferencias de zona horaria

Excel archivos no tienen una ubicación o zona horaria inherentes. Cada vez que un usuario abre el libro, su sesión usa la zona horaria local de ese usuario para los cálculos de fecha. Power Automate siempre usa UTC.

Si el script usa fechas u horas, puede haber diferencias de comportamiento cuando el script se prueba localmente frente a cuando se ejecuta a través de Power Automate. Power Automate permite convertir, dar formato y ajustar tiempos. Consulte [Trabajar](https://flow.microsoft.com/blog/working-with-dates-and-times/) con fechas y horas dentro de los flujos para obtener instrucciones sobre cómo usar esas funciones en Power Automate y [ `main` Parámetros:](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) pasar datos a un script para obtener información sobre cómo proporcionar esa información de hora para el script.

## <a name="script-parameter-fields-or-returned-output-not-appearing-in-power-automate"></a>Campos de parámetro script o resultados devueltos que no aparecen en Power Automate

Hay dos motivos por los que los parámetros o los datos devueltos de un script no se reflejan con precisión en el generador Power Automate flujo.

- La firma del script (los parámetros o el valor devuelto) ha cambiado desde que se Excel conector de empresa **(en** línea).
- La firma de script usa tipos no admitidos. Compruebe los tipos con [](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) las [](../develop/power-automate-integration.md#return-data-from-a-script) listas en los parámetros y devuelve secciones de [Ejecutar Office scripts con Power Automate](../develop/power-automate-integration.md) artículo.

La firma de un script se almacena con **el Excel empresarial (en línea)** cuando se crea. Quite el conector antiguo y cree uno nuevo para obtener los parámetros más recientes y devolver valores para la **acción Ejecutar script.**

## <a name="see-also"></a>Vea también

- [Solucionar problemas Office scripts](troubleshooting.md)
- [Ejecute Office scripts con Power Automate](../develop/power-automate-integration.md)
- [Excel Documentación de referencia del conector en línea (empresa)](/connectors/excelonlinebusiness/)
