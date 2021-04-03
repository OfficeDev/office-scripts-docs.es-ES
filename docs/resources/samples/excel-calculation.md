---
title: Administrar el modo de cálculo en Excel
description: Obtenga información sobre cómo usar scripts de Office para administrar el modo de cálculo en Excel en la web.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 0239437c7b52dca1fd8d1a4fc66bab7965cbd91a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571627"
---
# <a name="manage-calculation-mode-in-excel"></a>Administrar el modo de cálculo en Excel

En este ejemplo se muestra cómo usar el modo [de cálculo](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) y calcular métodos en Excel en la web mediante scripts de Office. Puede probar el script en cualquier archivo de Excel.

## <a name="scenario"></a>Escenario

En Excel en la web, el modo de cálculo de un archivo se puede controlar mediante programación mediante API. Las siguientes acciones son posibles mediante scripts de Office.

1. Obtener el modo de cálculo.
1. Establecer el modo de cálculo.
1. Calcular fórmulas de Excel para los archivos que se establecen en el modo manual (también denominado recalcular).

## <a name="sample-code-control-calculation-mode"></a>Código de ejemplo: Modo de cálculo de control

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set calculation mode.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Calculate (for manual mode files).
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a>Vídeo de aprendizaje: Administrar el modo de cálculo

[![Ver vídeo paso a paso sobre cómo administrar el modo de cálculo en Excel en la web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Vídeo paso a paso sobre cómo administrar el modo de cálculo en Excel en la web")
