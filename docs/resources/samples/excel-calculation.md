---
title: Administrar el modo de cálculo en Excel
description: Obtenga información sobre cómo usar Office scripts para administrar el modo de cálculo en Excel en la Web.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 34a14874197ffda8487df5e450e3dcab980f7ed5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232455"
---
# <a name="manage-calculation-mode-in-excel"></a>Administrar el modo de cálculo en Excel

En este ejemplo se muestra cómo usar el modo [de cálculo](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) y calcular métodos en Excel en la Web usar Office scripts. Puede probar el script en cualquier archivo Excel archivo.

## <a name="scenario"></a>Escenario

En Excel en la Web, el modo de cálculo de un archivo se puede controlar mediante programación mediante API. Las siguientes acciones son posibles mediante Office scripts.

1. Obtener el modo de cálculo.
1. Establecer el modo de cálculo.
1. Calcule Excel fórmulas de archivos establecidos en el modo manual (también denominado recalcular).

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

[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/iw6O8QH01CI).
