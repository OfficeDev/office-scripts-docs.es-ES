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
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="2da4b-103">Administrar el modo de cálculo en Excel</span><span class="sxs-lookup"><span data-stu-id="2da4b-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="2da4b-104">En este ejemplo se muestra cómo usar el modo [de cálculo](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) y calcular métodos en Excel en la Web usar Office scripts.</span><span class="sxs-lookup"><span data-stu-id="2da4b-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="2da4b-105">Puede probar el script en cualquier archivo Excel archivo.</span><span class="sxs-lookup"><span data-stu-id="2da4b-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="2da4b-106">Escenario</span><span class="sxs-lookup"><span data-stu-id="2da4b-106">Scenario</span></span>

<span data-ttu-id="2da4b-107">Los libros con un gran número de fórmulas pueden tardar un tiempo en volver a calcularse.</span><span class="sxs-lookup"><span data-stu-id="2da4b-107">Workbooks with large numbers of formulas can take a while to recalculate.</span></span> <span data-ttu-id="2da4b-108">En lugar de Excel control cuando se realiza un cálculo, puede administrarlos como parte del script.</span><span class="sxs-lookup"><span data-stu-id="2da4b-108">Rather than letting Excel control when calculations happen, you can manage them as part of your script.</span></span> <span data-ttu-id="2da4b-109">Esto ayudará con el rendimiento en determinados escenarios.</span><span class="sxs-lookup"><span data-stu-id="2da4b-109">This will help with performance in certain scenarios.</span></span>

<span data-ttu-id="2da4b-110">El script de ejemplo establece el modo de cálculo en manual.</span><span class="sxs-lookup"><span data-stu-id="2da4b-110">The sample script sets the calculation mode to manual.</span></span> <span data-ttu-id="2da4b-111">Esto significa que el libro solo recalculará fórmulas cuando el script lo indique (o calcule manualmente a través [de la interfaz de usuario](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)).</span><span class="sxs-lookup"><span data-stu-id="2da4b-111">This means that the workbook will only recalculate formulas when the script tells it to (or you [manually calculate through the UI](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)).</span></span> <span data-ttu-id="2da4b-112">A continuación, el script muestra el modo de cálculo actual y vuelve a calcular completamente todo el libro.</span><span class="sxs-lookup"><span data-stu-id="2da4b-112">The script then displays the current calculation mode and fully recalculates the entire workbook.</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="2da4b-113">Código de ejemplo: Modo de cálculo de control</span><span class="sxs-lookup"><span data-stu-id="2da4b-113">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="2da4b-114">Vídeo de aprendizaje: Administrar el modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="2da4b-114">Training video: Manage calculation mode</span></span>

<span data-ttu-id="2da4b-115">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="2da4b-115">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
