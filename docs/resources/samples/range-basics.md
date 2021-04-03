---
title: Conceptos básicos del intervalo en scripts de Office
description: Obtenga información básica sobre cómo usar el objeto Range en scripts de Office.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 73eeba086aace6262c624de9074ffb301f6532bd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571552"
---
# <a name="range-basics"></a><span data-ttu-id="42e56-103">Conceptos básicos del intervalo</span><span class="sxs-lookup"><span data-stu-id="42e56-103">Range basics</span></span>

<span data-ttu-id="42e56-104">`Range` es el objeto fundamental dentro del modelo de objetos de Excel de Scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="42e56-104">`Range` is the foundational object within the Office Scripts Excel object model.</span></span> <span data-ttu-id="42e56-105">[Las API de rango](/javascript/api/office-scripts/excelscript/excelscript.range) permiten el acceso a los datos y el formato disponibles en la cuadrícula y vinculan otros objetos clave dentro de Excel, como hojas de cálculo, tablas, gráficos, etc.</span><span class="sxs-lookup"><span data-stu-id="42e56-105">[Range APIs](/javascript/api/office-scripts/excelscript/excelscript.range) allow access to both data and format available on the grid and link other key objects within Excel such as worksheets, tables, charts, etc.</span></span>

<span data-ttu-id="42e56-106">Un intervalo se identifica con su dirección como "A1:B4" o con un elemento con nombre, que es una clave con nombre para un conjunto determinado de celdas.</span><span class="sxs-lookup"><span data-stu-id="42e56-106">A range is identified using its address such as "A1:B4" or using a named-item, which is a named key for a given set of cells.</span></span> <span data-ttu-id="42e56-107">En el modelo de objetos de Excel, una celda y un grupo de celdas se _denominan rango_.</span><span class="sxs-lookup"><span data-stu-id="42e56-107">In the Excel object model, both a cell and group of cells are referred as _range_.</span></span> <span data-ttu-id="42e56-108">`Range` puede contener atributos de nivel de celda, como datos dentro de una celda y también atributos de celda y de nivel de celda, como formato, bordes, etc.</span><span class="sxs-lookup"><span data-stu-id="42e56-108">`Range` can contain cell-level attributes such as data within a cell and also cell and cells-level attributes such as format, borders, etc.</span></span>

<span data-ttu-id="42e56-109">`Range` también se puede obtener a través de la selección del usuario que consta de al menos una celda.</span><span class="sxs-lookup"><span data-stu-id="42e56-109">`Range` can also be obtained via user's selection that consists of at least one cell.</span></span> <span data-ttu-id="42e56-110">A medida que interactúa con el rango, es importante mantener estas relaciones de celda y rango claras.</span><span class="sxs-lookup"><span data-stu-id="42e56-110">As you interact with the range, it's important to keep these cell and range relationships clear.</span></span>

<span data-ttu-id="42e56-111">A continuación se muestra el conjunto principal de getters, establecedores y otros métodos útiles más usados en scripts.</span><span class="sxs-lookup"><span data-stu-id="42e56-111">Following are the core set of getters, setters, and other useful methods most often used in scripts.</span></span> <span data-ttu-id="42e56-112">Este es un buen punto de partida para el recorrido de la API.</span><span class="sxs-lookup"><span data-stu-id="42e56-112">This is a great starting point for your API journey.</span></span> <span data-ttu-id="42e56-113">Las secciones posteriores agrupan los métodos y ayudan a crear un modelo mental a medida que empiezas a desbloquear las `Range` API del objeto.</span><span class="sxs-lookup"><span data-stu-id="42e56-113">The later sections group the methods and help to build a mental model as you begin to unlock the `Range` object's APIs.</span></span>

## <a name="example-scripts"></a><span data-ttu-id="42e56-114">Scripts de ejemplo</span><span class="sxs-lookup"><span data-stu-id="42e56-114">Example scripts</span></span>

* [<span data-ttu-id="42e56-115">Lectura y escritura básicas</span><span class="sxs-lookup"><span data-stu-id="42e56-115">Basic read and write</span></span>](#basic-read-and-write)
* [<span data-ttu-id="42e56-116">Agregar fila al final de la hoja de cálculo</span><span class="sxs-lookup"><span data-stu-id="42e56-116">Add row at the end of worksheet</span></span>](#add-row-at-the-end-of-worksheet)
* [<span data-ttu-id="42e56-117">Borrar filtro de columna</span><span class="sxs-lookup"><span data-stu-id="42e56-117">Clear column filter</span></span>](clear-table-filter-for-active-cell.md)
* [<span data-ttu-id="42e56-118">Color de cada celda con color único</span><span class="sxs-lookup"><span data-stu-id="42e56-118">Color each cell with unique color</span></span>](#color-each-cell-with-unique-color)
* [<span data-ttu-id="42e56-119">Intervalo de actualización con valores mediante matriz 2D</span><span class="sxs-lookup"><span data-stu-id="42e56-119">Update range with values using 2-dimensional (2D) array</span></span>](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a><span data-ttu-id="42e56-120">Lectura y escritura básicas</span><span class="sxs-lookup"><span data-stu-id="42e56-120">Basic read and write</span></span>

```TypeScript
/**
 * This script demonstrates basic read-write operations on the Range object.
 */
function main(workbook: ExcelScript.Workbook) {
  const cell = workbook.getActiveCell();
  const prevValue = cell.getValue();
  if (prevValue) {
      console.log(`Active cell's value is: ${prevValue}`);
  } else {
      console.log("Setting active cell's value..");
      cell.setValue("Sample");
  }

  // Get cell next to the right column and set its value and fill color.
  const nextCell = cell.getOffsetRange(0,1);
  nextCell.setValue("Next cell");
  console.log(`Next cell's address is: ${nextCell.getAddress()}`);
  console.log("Setting fill color and font color of next cell...");
  nextCell.getFormat().getFill().setColor("Magenta");
  nextCell.getFormat().getFill().setColor("Cyan");

  // Get the target range address to update with 2-dimensional value.
  const dataRange = nextCell.getOffsetRange(1, 0).getResizedRange(2, 1);
  const DATA = [
    [10, 7],
    [8, 15],
    [12, 1]
  ];
  console.log(`Updating range ${dataRange.getAddress()} with values: ${DATA}`);
  dataRange.setValues(DATA);

  // Formula range.
  const formulaRange = dataRange.getOffsetRange(3, 0).getRow(0);
  console.log(`Updating formula for range: ${formulaRange.getAddress()}`)
  // Since relative formula is being set, we can set the formula of the entire range to the same value.
  formulaRange.setFormulaR1C1("=SUM(R[-3]C:R[-1]C)");
  console.log(`Updating number format for range: ${formulaRange.getAddress()}`)
  // Since the number format is common to the entire range, we can set it to a common format.
  formulaRange.setNumberFormat("0.00");
  return;
}
```

### <a name="add-row-at-the-end-of-worksheet"></a><span data-ttu-id="42e56-121">Agregar fila al final de la hoja de cálculo</span><span class="sxs-lookup"><span data-stu-id="42e56-121">Add row at the end of worksheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for the update.
    if (usedRange) {
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);
    targetRange.setValues([data]);
    return;
}
```

### <a name="color-each-cell-with-unique-color"></a><span data-ttu-id="42e56-122">Color de cada celda con color único</span><span class="sxs-lookup"><span data-stu-id="42e56-122">Color each cell with unique color</span></span>

```TypeScript
/**
 * This sample demonstrates how to iterate over a selected range and set cell property.
   It colors each cell within the selected range with a random color.
 */
function main(workbook: ExcelScript.Workbook) {

    const syncStart = new Date().getTime();
    // Get selected range
    const range = workbook.getSelectedRange();
    const rows = range.getRowCount();
    const cols = range.getColumnCount();
    console.log("Start");

    // Color each cell with random color.
    for (let row = 0; row < rows; row++) {
        for (let col = 0; col < cols; col++) {
            range
                .getCell(row, col)
                .getFormat()
                .getFill()
                .setColor(`#${Math.random().toString(16).substr(-6)}`);
        }
    }

    console.log("End");
    const syncEnd = new Date().getTime();
    console.log("Completed, took: " + (syncEnd - syncStart) / 1000 + " Sec");
}
```

### <a name="update-range-with-values-using-2d-array"></a><span data-ttu-id="42e56-123">Actualizar intervalo con valores mediante matriz 2D</span><span class="sxs-lookup"><span data-stu-id="42e56-123">Update range with values using 2D array</span></span>

<span data-ttu-id="42e56-124">Calcula dinámicamente la dimensión de rango que se actualizará en función de los valores de matriz 2D.</span><span class="sxs-lookup"><span data-stu-id="42e56-124">Dynamically calculates the range dimension to update based on 2D array values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const currentCell = workbook.getActiveCell();
  let inputRange = computeTargetRange(currentCell, DATA);
  // Set range values.
  console.log(inputRange.getAddress());
  inputRange.setValues(DATA);
  // Call a helper function to place border around the range.
  borderAround(inputRange);
}

/**
 * A helper function that computes the target range given the target range's starting cell and selected range. 
 */
function computeTargetRange(targetCell: ExcelScript.Range, data: string[][]): ExcelScript.Range {
  const targetRange = targetCell.getResizedRange(data.length - 1, data[0].length - 1);
  return targetRange;
}

/**
 * A helper function that places a border around the range.
 */
function borderAround(range: ExcelScript.Range): void {
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.dash);
  return;
}

// Values used for range setup.
const DATA = [
  ['Item', 'Bread', 'Donuts', 'Cookies', 'Cakes', 'Pies'],
  ['Amount', '2', '1.5', '4', '12', '26']
]
```

## <a name="training-videos-range-basics"></a><span data-ttu-id="42e56-125">Vídeos de aprendizaje: conceptos básicos del rango</span><span class="sxs-lookup"><span data-stu-id="42e56-125">Training videos: Range basics</span></span>

<span data-ttu-id="42e56-126">_Conceptos básicos del intervalo_</span><span class="sxs-lookup"><span data-stu-id="42e56-126">_Range basics_</span></span>

<span data-ttu-id="42e56-127">[![Ver vídeo paso a paso en conceptos básicos del intervalo](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Vídeo paso a paso sobre los conceptos básicos del intervalo")</span><span class="sxs-lookup"><span data-stu-id="42e56-127">[![Watch step-by-step video on Range basics](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Step-by-step video on Range basics")</span></span>

<span data-ttu-id="42e56-128">_Agregar fila al final de la hoja de cálculo_</span><span class="sxs-lookup"><span data-stu-id="42e56-128">_Add row at the end of worksheet_</span></span>

<span data-ttu-id="42e56-129">[![Ver vídeo paso a paso sobre cómo agregar una fila al final de una hoja de cálculo](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Vídeo paso a paso sobre cómo agregar una fila al final de una hoja de cálculo")</span><span class="sxs-lookup"><span data-stu-id="42e56-129">[![Watch step-by-step video on how to add a row at the end of a worksheet](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Step-by-step video on how to add a row at the end of a worksheet")</span></span>

## <a name="methods-that-return-some-range-metadata"></a><span data-ttu-id="42e56-130">Métodos que devuelven algunos metadatos de intervalo</span><span class="sxs-lookup"><span data-stu-id="42e56-130">Methods that return some range metadata</span></span>

* <span data-ttu-id="42e56-131">getAddress(), getAddressLocal()</span><span class="sxs-lookup"><span data-stu-id="42e56-131">getAddress(), getAddressLocal()</span></span>
* <span data-ttu-id="42e56-132">getCellCount()</span><span class="sxs-lookup"><span data-stu-id="42e56-132">getCellCount()</span></span>
* <span data-ttu-id="42e56-133">getRowCount(), getColumnCount()</span><span class="sxs-lookup"><span data-stu-id="42e56-133">getRowCount(), getColumnCount()</span></span>

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a><span data-ttu-id="42e56-134">Métodos que devuelven datos/constantes asociados a un intervalo determinado</span><span class="sxs-lookup"><span data-stu-id="42e56-134">Methods that return data/constants associated with a given range</span></span>

### <a name="returned-as-single-cell-value"></a><span data-ttu-id="42e56-135">Devuelto como valor de celda única</span><span class="sxs-lookup"><span data-stu-id="42e56-135">Returned as single cell value</span></span>

* <span data-ttu-id="42e56-136">getFormula(), getFormulaLocal()</span><span class="sxs-lookup"><span data-stu-id="42e56-136">getFormula(), getFormulaLocal()</span></span>
* <span data-ttu-id="42e56-137">getFormulaR1C1()</span><span class="sxs-lookup"><span data-stu-id="42e56-137">getFormulaR1C1()</span></span>
* <span data-ttu-id="42e56-138">getNumberFormat(), getNumberFormatLocal()</span><span class="sxs-lookup"><span data-stu-id="42e56-138">getNumberFormat(), getNumberFormatLocal()</span></span>
* <span data-ttu-id="42e56-139">getText()</span><span class="sxs-lookup"><span data-stu-id="42e56-139">getText()</span></span>
* <span data-ttu-id="42e56-140">getValue()</span><span class="sxs-lookup"><span data-stu-id="42e56-140">getValue()</span></span>
* <span data-ttu-id="42e56-141">getValueType()</span><span class="sxs-lookup"><span data-stu-id="42e56-141">getValueType()</span></span>

### <a name="returned-as-2d-arrays-whole-range"></a><span data-ttu-id="42e56-142">Devuelto como matrices 2D (intervalo completo)</span><span class="sxs-lookup"><span data-stu-id="42e56-142">Returned as 2D arrays (whole range)</span></span>

* <span data-ttu-id="42e56-143">getFormulas(), getFormulasLocal()</span><span class="sxs-lookup"><span data-stu-id="42e56-143">getFormulas(), getFormulasLocal()</span></span>
* <span data-ttu-id="42e56-144">getFormulasR1C1()</span><span class="sxs-lookup"><span data-stu-id="42e56-144">getFormulasR1C1()</span></span>
* <span data-ttu-id="42e56-145">getNumberFormatCategories()</span><span class="sxs-lookup"><span data-stu-id="42e56-145">getNumberFormatCategories()</span></span>
* <span data-ttu-id="42e56-146">getNumberFormats(), getNumberFormatsLocal()</span><span class="sxs-lookup"><span data-stu-id="42e56-146">getNumberFormats(), getNumberFormatsLocal()</span></span>
* <span data-ttu-id="42e56-147">getTexts()</span><span class="sxs-lookup"><span data-stu-id="42e56-147">getTexts()</span></span>
* <span data-ttu-id="42e56-148">getValues()</span><span class="sxs-lookup"><span data-stu-id="42e56-148">getValues()</span></span>
* <span data-ttu-id="42e56-149">getValueTypes()</span><span class="sxs-lookup"><span data-stu-id="42e56-149">getValueTypes()</span></span>
* <span data-ttu-id="42e56-150">getHidden()</span><span class="sxs-lookup"><span data-stu-id="42e56-150">getHidden()</span></span>
* <span data-ttu-id="42e56-151">getIsEntireRow()</span><span class="sxs-lookup"><span data-stu-id="42e56-151">getIsEntireRow()</span></span>
* <span data-ttu-id="42e56-152">getIsEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="42e56-152">getIsEntireColumn()</span></span>

## <a name="methods-that-return-other-range-object"></a><span data-ttu-id="42e56-153">Métodos que devuelven otro objeto range</span><span class="sxs-lookup"><span data-stu-id="42e56-153">Methods that return other range object</span></span>

* <span data-ttu-id="42e56-154">getSurroundingRegion(): similar a CurrentRegion en VBA</span><span class="sxs-lookup"><span data-stu-id="42e56-154">getSurroundingRegion() -- similar to CurrentRegion in VBA</span></span>
* <span data-ttu-id="42e56-155">getCell(row, column)</span><span class="sxs-lookup"><span data-stu-id="42e56-155">getCell(row, column)</span></span>
* <span data-ttu-id="42e56-156">getColumn(column)</span><span class="sxs-lookup"><span data-stu-id="42e56-156">getColumn(column)</span></span>
* <span data-ttu-id="42e56-157">getColumnHidden()</span><span class="sxs-lookup"><span data-stu-id="42e56-157">getColumnHidden()</span></span>
* <span data-ttu-id="42e56-158">getColumnsAfter(count)</span><span class="sxs-lookup"><span data-stu-id="42e56-158">getColumnsAfter(count)</span></span>
* <span data-ttu-id="42e56-159">getColumnsBefore(count)</span><span class="sxs-lookup"><span data-stu-id="42e56-159">getColumnsBefore(count)</span></span>
* <span data-ttu-id="42e56-160">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="42e56-160">getEntireColumn()</span></span>
* <span data-ttu-id="42e56-161">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="42e56-161">getEntireRow()</span></span>
* <span data-ttu-id="42e56-162">getLastCell()</span><span class="sxs-lookup"><span data-stu-id="42e56-162">getLastCell()</span></span>
* <span data-ttu-id="42e56-163">getLastColumn()</span><span class="sxs-lookup"><span data-stu-id="42e56-163">getLastColumn()</span></span>
* <span data-ttu-id="42e56-164">getLastRow()</span><span class="sxs-lookup"><span data-stu-id="42e56-164">getLastRow()</span></span>
* <span data-ttu-id="42e56-165">getRow(row)</span><span class="sxs-lookup"><span data-stu-id="42e56-165">getRow(row)</span></span>
* <span data-ttu-id="42e56-166">getRowHidden()</span><span class="sxs-lookup"><span data-stu-id="42e56-166">getRowHidden()</span></span>
* <span data-ttu-id="42e56-167">getRowsAbove(count)</span><span class="sxs-lookup"><span data-stu-id="42e56-167">getRowsAbove(count)</span></span>
* <span data-ttu-id="42e56-168">getRowsBelow(count)</span><span class="sxs-lookup"><span data-stu-id="42e56-168">getRowsBelow(count)</span></span>

<span data-ttu-id="42e56-169">**Importante/Interesante**</span><span class="sxs-lookup"><span data-stu-id="42e56-169">**Important/Interesting**</span></span>

* <span data-ttu-id="42e56-170">_workbook_.getSelectedRange()</span><span class="sxs-lookup"><span data-stu-id="42e56-170">_workbook_.getSelectedRange()</span></span>
* <span data-ttu-id="42e56-171">_workbook_.getActiveCell()</span><span class="sxs-lookup"><span data-stu-id="42e56-171">_workbook_.getActiveCell()</span></span>
* <span data-ttu-id="42e56-172">getUsedRange(valuesOnly)</span><span class="sxs-lookup"><span data-stu-id="42e56-172">getUsedRange(valuesOnly)</span></span>
* <span data-ttu-id="42e56-173">getAbsoluteResizedRange(numRows, numColumns)</span><span class="sxs-lookup"><span data-stu-id="42e56-173">getAbsoluteResizedRange(numRows, numColumns)</span></span>
* <span data-ttu-id="42e56-174">getOffsetRange(rowOffset, columnOffset)</span><span class="sxs-lookup"><span data-stu-id="42e56-174">getOffsetRange(rowOffset, columnOffset)</span></span>
* <span data-ttu-id="42e56-175">getResizedRange(deltaRows, deltaColumns)</span><span class="sxs-lookup"><span data-stu-id="42e56-175">getResizedRange(deltaRows, deltaColumns)</span></span>

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a><span data-ttu-id="42e56-176">Métodos que devuelven un objeto range en relación con otro objeto range</span><span class="sxs-lookup"><span data-stu-id="42e56-176">Methods that return a range object in relation to another range object</span></span>

* <span data-ttu-id="42e56-177">getBoundingRect(anotherRange)</span><span class="sxs-lookup"><span data-stu-id="42e56-177">getBoundingRect(anotherRange)</span></span>
* <span data-ttu-id="42e56-178">getIntersection(anotherRange)</span><span class="sxs-lookup"><span data-stu-id="42e56-178">getIntersection(anotherRange)</span></span>

## <a name="methods-that-return-other-objects-non-range-objects"></a><span data-ttu-id="42e56-179">Métodos que devuelven otros objetos (objetos que no son de intervalo)</span><span class="sxs-lookup"><span data-stu-id="42e56-179">Methods that return other objects (non-range objects)</span></span>

* <span data-ttu-id="42e56-180">getDirectPrecedents()</span><span class="sxs-lookup"><span data-stu-id="42e56-180">getDirectPrecedents()</span></span>
* <span data-ttu-id="42e56-181">getWorksheet()</span><span class="sxs-lookup"><span data-stu-id="42e56-181">getWorksheet()</span></span>
* <span data-ttu-id="42e56-182">getTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="42e56-182">getTables(fullyContained)</span></span>
* <span data-ttu-id="42e56-183">getPivotTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="42e56-183">getPivotTables(fullyContained)</span></span>
* <span data-ttu-id="42e56-184">getDataValidation()</span><span class="sxs-lookup"><span data-stu-id="42e56-184">getDataValidation()</span></span>
* <span data-ttu-id="42e56-185">getPredefinedCellStyle()</span><span class="sxs-lookup"><span data-stu-id="42e56-185">getPredefinedCellStyle()</span></span>

## <a name="set-methods"></a><span data-ttu-id="42e56-186">Establecer métodos</span><span class="sxs-lookup"><span data-stu-id="42e56-186">Set methods</span></span>

### <a name="singular-cell-set-methods"></a><span data-ttu-id="42e56-187">Métodos de conjunto de celdas singulares</span><span class="sxs-lookup"><span data-stu-id="42e56-187">Singular cell set methods</span></span>

* <span data-ttu-id="42e56-188">setFormula(formula)</span><span class="sxs-lookup"><span data-stu-id="42e56-188">setFormula(formula)</span></span>
* <span data-ttu-id="42e56-189">setFormulaLocal(formulaLocal)</span><span class="sxs-lookup"><span data-stu-id="42e56-189">setFormulaLocal(formulaLocal)</span></span>
* <span data-ttu-id="42e56-190">setFormulaR1C1(formulaR1C1)</span><span class="sxs-lookup"><span data-stu-id="42e56-190">setFormulaR1C1(formulaR1C1)</span></span>
* <span data-ttu-id="42e56-191">setNumberFormatLocal(numberFormatLocal)</span><span class="sxs-lookup"><span data-stu-id="42e56-191">setNumberFormatLocal(numberFormatLocal)</span></span>
* <span data-ttu-id="42e56-192">setValue(value)</span><span class="sxs-lookup"><span data-stu-id="42e56-192">setValue(value)</span></span>

### <a name="2d--entire-range-set-methods"></a><span data-ttu-id="42e56-193">Métodos 2D/conjunto de intervalos completos</span><span class="sxs-lookup"><span data-stu-id="42e56-193">2D / entire range set methods</span></span>

* <span data-ttu-id="42e56-194">setFormulas(formulas)</span><span class="sxs-lookup"><span data-stu-id="42e56-194">setFormulas(formulas)</span></span>
* <span data-ttu-id="42e56-195">setFormulasLocal(formulasLocal)</span><span class="sxs-lookup"><span data-stu-id="42e56-195">setFormulasLocal(formulasLocal)</span></span>
* <span data-ttu-id="42e56-196">setFormulasR1C1(formulasR1C1)</span><span class="sxs-lookup"><span data-stu-id="42e56-196">setFormulasR1C1(formulasR1C1)</span></span>
* <span data-ttu-id="42e56-197">setNumberFormat(numberFormat)</span><span class="sxs-lookup"><span data-stu-id="42e56-197">setNumberFormat(numberFormat)</span></span>
* <span data-ttu-id="42e56-198">setNumberFormats(numberFormats)</span><span class="sxs-lookup"><span data-stu-id="42e56-198">setNumberFormats(numberFormats)</span></span>
* <span data-ttu-id="42e56-199">setNumberFormatsLocal(numberFormatsLocal)</span><span class="sxs-lookup"><span data-stu-id="42e56-199">setNumberFormatsLocal(numberFormatsLocal)</span></span>
* <span data-ttu-id="42e56-200">setValues(values)</span><span class="sxs-lookup"><span data-stu-id="42e56-200">setValues(values)</span></span>

## <a name="other-methods"></a><span data-ttu-id="42e56-201">Otros métodos</span><span class="sxs-lookup"><span data-stu-id="42e56-201">Other methods</span></span>

* <span data-ttu-id="42e56-202">merge(across)</span><span class="sxs-lookup"><span data-stu-id="42e56-202">merge(across)</span></span>
* <span data-ttu-id="42e56-203">unmerge()</span><span class="sxs-lookup"><span data-stu-id="42e56-203">unmerge()</span></span>

## <a name="coming-soon"></a><span data-ttu-id="42e56-204">Próximamente</span><span class="sxs-lookup"><span data-stu-id="42e56-204">Coming soon</span></span>

* <span data-ttu-id="42e56-205">API perimetrales de intervalo</span><span class="sxs-lookup"><span data-stu-id="42e56-205">Range edge APIs</span></span>
