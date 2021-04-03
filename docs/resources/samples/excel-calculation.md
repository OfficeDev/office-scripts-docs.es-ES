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
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="b6ac0-103">Administrar el modo de cálculo en Excel</span><span class="sxs-lookup"><span data-stu-id="b6ac0-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="b6ac0-104">En este ejemplo se muestra cómo usar el modo [de cálculo](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) y calcular métodos en Excel en la web mediante scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="b6ac0-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="b6ac0-105">Puede probar el script en cualquier archivo de Excel.</span><span class="sxs-lookup"><span data-stu-id="b6ac0-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="b6ac0-106">Escenario</span><span class="sxs-lookup"><span data-stu-id="b6ac0-106">Scenario</span></span>

<span data-ttu-id="b6ac0-107">En Excel en la web, el modo de cálculo de un archivo se puede controlar mediante programación mediante API.</span><span class="sxs-lookup"><span data-stu-id="b6ac0-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="b6ac0-108">Las siguientes acciones son posibles mediante scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="b6ac0-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="b6ac0-109">Obtener el modo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="b6ac0-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="b6ac0-110">Establecer el modo de cálculo.</span><span class="sxs-lookup"><span data-stu-id="b6ac0-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="b6ac0-111">Calcular fórmulas de Excel para los archivos que se establecen en el modo manual (también denominado recalcular).</span><span class="sxs-lookup"><span data-stu-id="b6ac0-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="b6ac0-112">Código de ejemplo: Modo de cálculo de control</span><span class="sxs-lookup"><span data-stu-id="b6ac0-112">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="b6ac0-113">Vídeo de aprendizaje: Administrar el modo de cálculo</span><span class="sxs-lookup"><span data-stu-id="b6ac0-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="b6ac0-114">[![Ver vídeo paso a paso sobre cómo administrar el modo de cálculo en Excel en la web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Vídeo paso a paso sobre cómo administrar el modo de cálculo en Excel en la web")</span><span class="sxs-lookup"><span data-stu-id="b6ac0-114">[![Watch step-by-step video on how to manage calculation mode in Excel on the web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Step-by-step video on how to manage calculation mode in Excel on the web")</span></span>
