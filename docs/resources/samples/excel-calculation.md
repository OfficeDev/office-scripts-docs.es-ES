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
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="495de-103">Administrar el modo de cálculo en Excel</span><span class="sxs-lookup"><span data-stu-id="495de-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="495de-104">En este ejemplo se muestra cómo usar el modo [de cálculo](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) y calcular métodos en Excel en la Web usar Office scripts.</span><span class="sxs-lookup"><span data-stu-id="495de-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="495de-105">Puede probar el script en cualquier archivo Excel archivo.</span><span class="sxs-lookup"><span data-stu-id="495de-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="495de-106">Escenario</span><span class="sxs-lookup"><span data-stu-id="495de-106">Scenario</span></span>

<span data-ttu-id="495de-107">En Excel en la Web, el modo de cálculo de un archivo se puede controlar mediante programación mediante API.</span><span class="sxs-lookup"><span data-stu-id="495de-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="495de-108">Las siguientes acciones son posibles mediante Office scripts.</span><span class="sxs-lookup"><span data-stu-id="495de-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="495de-109">Obtener el modo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="495de-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="495de-110">Establecer el modo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="495de-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="495de-111">Calcule Excel fórmulas de archivos establecidos en el modo manual (también denominado recalcular).</span><span class="sxs-lookup"><span data-stu-id="495de-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="495de-112">Código de ejemplo: Modo de cálculo de control</span><span class="sxs-lookup"><span data-stu-id="495de-112">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="495de-113">Vídeo de aprendizaje: Administrar el modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="495de-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="495de-114">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="495de-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
