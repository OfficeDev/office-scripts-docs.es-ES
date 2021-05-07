---
title: Generar un identificador único en un libro
description: Obtenga información sobre cómo usar Office scripts para generar un identificador único y agregar una fila a una tabla y un intervalo.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 62c930bfc638dc46b36daf81b6d1ec976c90a8d0
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232483"
---
# <a name="generate-a-unique-identifier-in-a-workbook"></a><span data-ttu-id="56bf9-103">Generar un identificador único en un libro</span><span class="sxs-lookup"><span data-stu-id="56bf9-103">Generate a unique identifier in a workbook</span></span>

<span data-ttu-id="56bf9-104">Este escenario ayuda a un usuario a generar un número de documento único con un formato específico y agregarlo como una entrada a un rango o tabla.</span><span class="sxs-lookup"><span data-stu-id="56bf9-104">This scenario helps a user generate a unique document number with a specific format and add it as an entry to a range or table.</span></span> <span data-ttu-id="56bf9-105">La nueva entrada o fila agregada contendrá el número de documento único recién generado y otros atributos pasados al script.</span><span class="sxs-lookup"><span data-stu-id="56bf9-105">The new entry or row added will contain the newly generated unique document number and a few other attributes passed to the script.</span></span>

<span data-ttu-id="56bf9-106">Hay dos versiones del ejemplo para este escenario.</span><span class="sxs-lookup"><span data-stu-id="56bf9-106">There are two versions of the sample for this scenario.</span></span>

* [<span data-ttu-id="56bf9-107">Versión 1: Leer y agregar una fila a una hoja de cálculo que contiene un intervalo sin formato</span><span class="sxs-lookup"><span data-stu-id="56bf9-107">Version 1: Read and add a row to a worksheet containing plain range</span></span>](#sample-code-generate-key-and-add-row-to-range)

    <span data-ttu-id="56bf9-108">_Antes de agregar la nueva fila_</span><span class="sxs-lookup"><span data-stu-id="56bf9-108">_Before the new row is added_</span></span>

    :::image type="content" source="../../images/document-number-generator-range-before.png" alt-text="Hoja de cálculo que muestra un rango de datos antes de agregar una fila":::

    <span data-ttu-id="56bf9-110">_Después de agregar la nueva fila_</span><span class="sxs-lookup"><span data-stu-id="56bf9-110">_After the new row is added_</span></span>

    :::image type="content" source="../../images/document-number-generator-range-after.png" alt-text="Una hoja de cálculo que muestra un rango de datos después de agregar una fila":::

* [<span data-ttu-id="56bf9-112">Versión 2: Leer y agregar una fila a una tabla</span><span class="sxs-lookup"><span data-stu-id="56bf9-112">Version 2: Read and add a row to a table</span></span>](#sample-code-generate-key-and-add-row-to-table)

    <span data-ttu-id="56bf9-113">_Antes de agregar la nueva fila_</span><span class="sxs-lookup"><span data-stu-id="56bf9-113">_Before the new row is added_</span></span>

    :::image type="content" source="../../images/document-number-generator-table-before.png" alt-text="Hoja de cálculo que muestra una tabla antes de agregar una fila":::

    <span data-ttu-id="56bf9-115">_Después de agregar la nueva fila_</span><span class="sxs-lookup"><span data-stu-id="56bf9-115">_After the new row is added_</span></span>

    :::image type="content" source="../../images/document-number-generator-table-after.png" alt-text="Una hoja de cálculo que muestra una tabla después de agregar una fila":::

## <a name="sample-excel-file"></a><span data-ttu-id="56bf9-117">Archivo Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="56bf9-117">Sample Excel file</span></span>

<span data-ttu-id="56bf9-118">Descargue el archivo <a href="document-number-generator.xlsx">document-number-generator.xlsx</a> se usa en esta solución para probarlo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="56bf9-118">Download the file <a href="document-number-generator.xlsx">document-number-generator.xlsx</a> used in this solution to try it out yourself!</span></span>

## <a name="sample-code-generate-key-and-add-row-to-range"></a><span data-ttu-id="56bf9-119">Código de ejemplo: Generar clave y agregar fila al intervalo</span><span class="sxs-lookup"><span data-stu-id="56bf9-119">Sample code: Generate key and add row to range</span></span>

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

## <a name="sample-code-generate-key-and-add-row-to-table"></a><span data-ttu-id="56bf9-120">Código de ejemplo: Generar clave y agregar fila a tabla</span><span class="sxs-lookup"><span data-stu-id="56bf9-120">Sample code: Generate key and add row to table</span></span>

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
