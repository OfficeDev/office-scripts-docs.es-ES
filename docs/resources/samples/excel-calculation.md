---
title: Administrar el modo de cálculo en Excel
description: Obtenga información sobre cómo usar Office scripts para administrar el modo de cálculo en Excel en la Web.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: a60fddc91b3a8f124a44722d0d75e6e9f239351d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285916"
---
# <a name="manage-calculation-mode-in-excel"></a>Administrar el modo de cálculo en Excel

En este ejemplo se muestra cómo usar el modo [de cálculo](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) y calcular métodos en Excel en la Web usar Office scripts. Puede probar el script en cualquier archivo Excel archivo.

## <a name="scenario"></a>Escenario

Los libros con un gran número de fórmulas pueden tardar un tiempo en volver a calcularse. En lugar de Excel control cuando se realiza un cálculo, puede administrarlos como parte del script. Esto ayudará con el rendimiento en determinados escenarios.

El script de ejemplo establece el modo de cálculo en manual. Esto significa que el libro solo recalculará fórmulas cuando el script lo indique (o calcule manualmente a través [de la interfaz de usuario](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)). A continuación, el script muestra el modo de cálculo actual y vuelve a calcular completamente todo el libro.

## <a name="sample-code-control-calculation-mode"></a>Código de ejemplo: Modo de cálculo de control

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set the calculation mode to manual.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get and log the calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Manually calculate the file.
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a>Vídeo de aprendizaje: Administrar el modo de cálculo

[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/iw6O8QH01CI).
