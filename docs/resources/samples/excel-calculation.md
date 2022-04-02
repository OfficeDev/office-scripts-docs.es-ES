---
title: Administrar el modo de cálculo en Excel
description: Obtenga información sobre cómo usar Office scripts para administrar el modo de cálculo en Excel en la Web.
ms.date: 05/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: fec88c904d95bfdab1514d44921f7fb1c6e9dd35
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585516"
---
# <a name="manage-calculation-mode-in-excel"></a>Administrar el modo de cálculo en Excel

En este ejemplo se muestra cómo usar el modo [de cálculo](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) y calcular los métodos en Excel en la Web usar Office scripts. Puede probar el script en cualquier archivo Excel archivo.

## <a name="scenario"></a>Escenario

Los libros con un gran número de fórmulas pueden tardar un tiempo en volver a calcularse. En lugar de Excel control cuando se realiza un cálculo, puede administrarlos como parte del script. Esto ayudará con el rendimiento en determinados escenarios.

El script de ejemplo establece el modo de cálculo en manual. Esto significa que el libro solo recalculará fórmulas cuando el script lo indique (o [calcule manualmente a través de la interfaz de usuario](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4)). A continuación, el script muestra el modo de cálculo actual y vuelve a calcular completamente todo el libro.

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

[Vea el recorrido de Sudhi Ramamurthy a través de esta muestra en YouTube](https://youtu.be/iw6O8QH01CI).
