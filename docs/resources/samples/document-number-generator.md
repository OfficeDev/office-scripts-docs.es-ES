---
title: Generar un identificador único en un libro
description: Obtenga información sobre cómo usar Office scripts para generar un identificador único y agregar una fila a una tabla y un intervalo.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 62c930bfc638dc46b36daf81b6d1ec976c90a8d0
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/08/2021
ms.locfileid: "52286108"
---
# <a name="generate-a-unique-identifier-in-a-workbook"></a>Generar un identificador único en un libro

Este escenario ayuda a un usuario a generar un número de documento único con un formato específico y agregarlo como una entrada a un rango o tabla. La nueva entrada o fila agregada contendrá el número de documento único recién generado y otros atributos pasados al script.

Hay dos versiones del ejemplo para este escenario.

* [Versión 1: Leer y agregar una fila a una hoja de cálculo que contiene un intervalo sin formato](#sample-code-generate-key-and-add-row-to-range)

    _Antes de agregar la nueva fila_

    :::image type="content" source="../../images/document-number-generator-range-before.png" alt-text="Hoja de cálculo que muestra un rango de datos antes de agregar una fila":::

    _Después de agregar la nueva fila_

    :::image type="content" source="../../images/document-number-generator-range-after.png" alt-text="Una hoja de cálculo que muestra un rango de datos después de agregar una fila":::

* [Versión 2: Leer y agregar una fila a una tabla](#sample-code-generate-key-and-add-row-to-table)

    _Antes de agregar la nueva fila_

    :::image type="content" source="../../images/document-number-generator-table-before.png" alt-text="Hoja de cálculo que muestra una tabla antes de agregar una fila":::

    _Después de agregar la nueva fila_

    :::image type="content" source="../../images/document-number-generator-table-after.png" alt-text="Una hoja de cálculo que muestra una tabla después de agregar una fila":::

## <a name="sample-excel-file"></a>Archivo Excel ejemplo

Descargue el archivo <a href="document-number-generator.xlsx">document-number-generator.xlsx</a> se usa en esta solución para probarlo usted mismo.

## <a name="sample-code-generate-key-and-add-row-to-range"></a>Código de ejemplo: Generar clave y agregar fila al intervalo

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // Object to hold key prefixes for each document type.
    const PREFIX  = {
        form: 'F',
        'work instruction': 'W'
    }

    // Length of the numeric part of the key.
    const KEYLENGTH = 6;

    // Parse the incoming string as object.
    const input:RequestData = JSON.parse(inputString);

    // Reject invalid request.
    if (input.docType.toLowerCase() !== 'form' && 
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get existing data in the sheet.
    const sheet = workbook.getWorksheet('PlainSheet'); /* plain range sheet */
    const range = sheet.getUsedRange();

    const data = range.getValues() as string[][];

    // Filter rows to match the incoming type and then extract the document number column (index 0) and then sort it. 
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the max document ID for the type.
    const maxId = selectIds[selectIds.length-1];

    // Extract numeric part.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the max key value, throw an error.
    if (nextNum >= (10 ** KEYLENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }
    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;
    
    // Compute next key value.
    const nextKey = prefixVal + '0'.repeat(KEYLENGTH).substring(0, KEYLENGTH - String(nextNum).length) + String(nextNum);
    
    // Get last row and compute next row address.
    const last = range.getLastRow();
    const target = last.getOffsetRange(1, 0);

    // Add a row with incoming data plus the computed key value.
    target.setValues([
      [
        nextKey, 
        /* Capitalize the document type. */
        input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
        input.documentName
      ]
    ])
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`)
    // Return the key value recorded in Excel.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```

## <a name="sample-code-generate-key-and-add-row-to-table"></a>Código de ejemplo: Generar clave y agregar fila a tabla

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // Object to hold key prefixes for each document type.
    const PREFIX = {
        form: 'F',
        'work instruction': 'W'
    }

    // Length of the numeric part of the key.
    const KEYLENGTH = 6;

    // Parse the incoming string as object.
    const input: RequestData = JSON.parse(inputString);

    // Reject invalid request.
    if (input.docType.toLowerCase() !== 'form' &&
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get existing data in the sheet.
    const sheet = workbook.getWorksheet('TableSheet'); /* table sheet */
    const table = sheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const data = range.getValues() as string[][];

    // Filter rows to match the incoming type and then extract the document number column (index 0) and then sort it.
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the max document ID for the type.
    const maxId = selectIds[selectIds.length - 1];


    // Extract numeric part.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the max key value, throw an error.
    if (nextNum >= (10 ** KEYLENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }
    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;

    // Compute next key value.
    const nextKey = prefixVal + '0'.repeat(KEYLENGTH).substring(0, KEYLENGTH - String(nextNum).length) + String(nextNum);

    // Add a row with incoming data plus the computed key value.
    table.addRow(-1, [
            nextKey,
            /* Capitalize the document type. */
            input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
            input.documentName
        ]);
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`)
    // Return the key value recorded in Excel.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```
