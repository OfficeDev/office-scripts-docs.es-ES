---
title: Agregar imágenes a un libro
description: Obtén información sobre cómo usar scripts de Office para agregar una imagen a un libro de trabajo y copiarla en todas las hojas.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 99c3cc2cacf6e535bdb882bb8414d23fd105be35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546040"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="016da-103">Agregar imágenes a un libro</span><span class="sxs-lookup"><span data-stu-id="016da-103">Add images to a workbook</span></span>

<span data-ttu-id="016da-104">En este ejemplo se muestra cómo trabajar con imágenes mediante un script de Office en Excel.</span><span class="sxs-lookup"><span data-stu-id="016da-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="016da-105">Escenario</span><span class="sxs-lookup"><span data-stu-id="016da-105">Scenario</span></span>

<span data-ttu-id="016da-106">Las imágenes ayudan con el branding, la identidad visual y las plantillas.</span><span class="sxs-lookup"><span data-stu-id="016da-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="016da-107">Ayudan a hacer un libro de trabajo más que una mesa gigante.</span><span class="sxs-lookup"><span data-stu-id="016da-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="016da-108">El primer ejemplo copia una imagen de una hoja de cálculo a otra.</span><span class="sxs-lookup"><span data-stu-id="016da-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="016da-109">Esto podría usarse para colocar el logotipo de su empresa en la misma posición en cada hoja.</span><span class="sxs-lookup"><span data-stu-id="016da-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="016da-110">En el segundo ejemplo se copia una imagen de una dirección URL.</span><span class="sxs-lookup"><span data-stu-id="016da-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="016da-111">Esto podría usarse para copiar fotos que un colega almacenó en una carpeta compartida a un libro de trabajo relacionado.</span><span class="sxs-lookup"><span data-stu-id="016da-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="016da-112">Archivo de Excel de ejemplo</span><span class="sxs-lookup"><span data-stu-id="016da-112">Sample Excel file</span></span>

<span data-ttu-id="016da-113">Descargar el archivo <a href="add-images.xlsx">add-images.xlsx</a> utilizado en estas muestras y probarlo usted mismo!</span><span class="sxs-lookup"><span data-stu-id="016da-113">Download the file <a href="add-images.xlsx">add-images.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="016da-114">Código de ejemplo: copie una imagen en las hojas de trabajo</span><span class="sxs-lookup"><span data-stu-id="016da-114">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="016da-115">Código de ejemplo: agregue una imagen de una dirección URL a un libro</span><span class="sxs-lookup"><span data-stu-id="016da-115">Sample code: Add an image from a URL to a workbook</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image)
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) 
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```