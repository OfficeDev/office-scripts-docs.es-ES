---
title: Agregar imágenes a un libro
description: Obtenga información sobre cómo usar scripts de Office para agregar una imagen a un libro y copiarla entre hojas.
ms.date: 07/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 78c7779cf4d524ed62bf8d419135863228b23d33
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572608"
---
# <a name="add-images-to-a-workbook"></a>Agregar imágenes a un libro

En este ejemplo se muestra cómo trabajar con imágenes mediante un script de Office en Excel.

## <a name="scenario"></a>Escenario

Las imágenes ayudan con la personalización de marca, la identidad visual y las plantillas. Ayudan a crear un libro más que una mesa gigante.

El primer ejemplo copia una imagen de una hoja de cálculo a otra. Esto podría usarse para colocar el logotipo de su empresa en la misma posición en cada hoja.

En el segundo ejemplo se copia una imagen de una dirección URL. Esto se podría usar para copiar las fotos que un compañero almacenó en una carpeta compartida en un libro relacionado.

## <a name="sample-excel-file"></a>Archivo de Excel de ejemplo

Descargue [add-images.xlsx](add-images.xlsx) de un libro listo para usar. Agregue los siguientes scripts y pruebe el ejemplo usted mismo.

## <a name="sample-code-copy-an-image-across-worksheets"></a>Código de ejemplo: Copia de una imagen entre hojas de cálculo

```TypeScript
/**
 * This script transfers an image from one worksheet to another.
 */
function main(workbook: ExcelScript.Workbook)
{
  // Get the worksheet with the image on it.
  let firstWorksheet = workbook.getWorksheet("FirstSheet");

  // Get the first image from the worksheet.
  // If a script added the image, you could add a name to make it easier to find.
  let image: ExcelScript.Image;
  firstWorksheet.getShapes().forEach((shape, index) => {
    if (shape.getType() === ExcelScript.ShapeType.image) {
      image = shape.getImage();
      return;
    }
  });

  // Copy the image to another worksheet.
  image.getShape().copyTo("SecondSheet");
}
```

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>Código de ejemplo: Agregar una imagen de una dirección URL a un libro

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://raw.githubusercontent.com/OfficeDev/office-scripts-docs/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image);
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) as string[];
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
