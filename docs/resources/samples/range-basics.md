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
# <a name="range-basics"></a>Conceptos básicos del intervalo

`Range` es el objeto fundamental dentro del modelo de objetos de Excel de Scripts de Office. [Las API de rango](/javascript/api/office-scripts/excelscript/excelscript.range) permiten el acceso a los datos y el formato disponibles en la cuadrícula y vinculan otros objetos clave dentro de Excel, como hojas de cálculo, tablas, gráficos, etc.

Un intervalo se identifica con su dirección como "A1:B4" o con un elemento con nombre, que es una clave con nombre para un conjunto determinado de celdas. En el modelo de objetos de Excel, una celda y un grupo de celdas se _denominan rango_. `Range` puede contener atributos de nivel de celda, como datos dentro de una celda y también atributos de celda y de nivel de celda, como formato, bordes, etc.

`Range` también se puede obtener a través de la selección del usuario que consta de al menos una celda. A medida que interactúa con el rango, es importante mantener estas relaciones de celda y rango claras.

A continuación se muestra el conjunto principal de getters, establecedores y otros métodos útiles más usados en scripts. Este es un buen punto de partida para el recorrido de la API. Las secciones posteriores agrupan los métodos y ayudan a crear un modelo mental a medida que empiezas a desbloquear las `Range` API del objeto.

## <a name="example-scripts"></a>Scripts de ejemplo

* [Lectura y escritura básicas](#basic-read-and-write)
* [Agregar fila al final de la hoja de cálculo](#add-row-at-the-end-of-worksheet)
* [Borrar filtro de columna](clear-table-filter-for-active-cell.md)
* [Color de cada celda con color único](#color-each-cell-with-unique-color)
* [Intervalo de actualización con valores mediante matriz 2D](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a>Lectura y escritura básicas

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

### <a name="add-row-at-the-end-of-worksheet"></a>Agregar fila al final de la hoja de cálculo

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

### <a name="color-each-cell-with-unique-color"></a>Color de cada celda con color único

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

### <a name="update-range-with-values-using-2d-array"></a>Actualizar intervalo con valores mediante matriz 2D

Calcula dinámicamente la dimensión de rango que se actualizará en función de los valores de matriz 2D.

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

## <a name="training-videos-range-basics"></a>Vídeos de aprendizaje: conceptos básicos del rango

_Conceptos básicos del intervalo_

[![Ver vídeo paso a paso en conceptos básicos del intervalo](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Vídeo paso a paso sobre los conceptos básicos del intervalo")

_Agregar fila al final de la hoja de cálculo_

[![Ver vídeo paso a paso sobre cómo agregar una fila al final de una hoja de cálculo](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Vídeo paso a paso sobre cómo agregar una fila al final de una hoja de cálculo")

## <a name="methods-that-return-some-range-metadata"></a>Métodos que devuelven algunos metadatos de intervalo

* getAddress(), getAddressLocal()
* getCellCount()
* getRowCount(), getColumnCount()

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a>Métodos que devuelven datos/constantes asociados a un intervalo determinado

### <a name="returned-as-single-cell-value"></a>Devuelto como valor de celda única

* getFormula(), getFormulaLocal()
* getFormulaR1C1()
* getNumberFormat(), getNumberFormatLocal()
* getText()
* getValue()
* getValueType()

### <a name="returned-as-2d-arrays-whole-range"></a>Devuelto como matrices 2D (intervalo completo)

* getFormulas(), getFormulasLocal()
* getFormulasR1C1()
* getNumberFormatCategories()
* getNumberFormats(), getNumberFormatsLocal()
* getTexts()
* getValues()
* getValueTypes()
* getHidden()
* getIsEntireRow()
* getIsEntireColumn()

## <a name="methods-that-return-other-range-object"></a>Métodos que devuelven otro objeto range

* getSurroundingRegion(): similar a CurrentRegion en VBA
* getCell(row, column)
* getColumn(column)
* getColumnHidden()
* getColumnsAfter(count)
* getColumnsBefore(count)
* getEntireColumn()
* getEntireRow()
* getLastCell()
* getLastColumn()
* getLastRow()
* getRow(row)
* getRowHidden()
* getRowsAbove(count)
* getRowsBelow(count)

**Importante/Interesante**

* _workbook_.getSelectedRange()
* _workbook_.getActiveCell()
* getUsedRange(valuesOnly)
* getAbsoluteResizedRange(numRows, numColumns)
* getOffsetRange(rowOffset, columnOffset)
* getResizedRange(deltaRows, deltaColumns)

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a>Métodos que devuelven un objeto range en relación con otro objeto range

* getBoundingRect(anotherRange)
* getIntersection(anotherRange)

## <a name="methods-that-return-other-objects-non-range-objects"></a>Métodos que devuelven otros objetos (objetos que no son de intervalo)

* getDirectPrecedents()
* getWorksheet()
* getTables(fullyContained)
* getPivotTables(fullyContained)
* getDataValidation()
* getPredefinedCellStyle()

## <a name="set-methods"></a>Establecer métodos

### <a name="singular-cell-set-methods"></a>Métodos de conjunto de celdas singulares

* setFormula(formula)
* setFormulaLocal(formulaLocal)
* setFormulaR1C1(formulaR1C1)
* setNumberFormatLocal(numberFormatLocal)
* setValue(value)

### <a name="2d--entire-range-set-methods"></a>Métodos 2D/conjunto de intervalos completos

* setFormulas(formulas)
* setFormulasLocal(formulasLocal)
* setFormulasR1C1(formulasR1C1)
* setNumberFormat(numberFormat)
* setNumberFormats(numberFormats)
* setNumberFormatsLocal(numberFormatsLocal)
* setValues(values)

## <a name="other-methods"></a>Otros métodos

* merge(across)
* unmerge()

## <a name="coming-soon"></a>Próximamente

* API perimetrales de intervalo
