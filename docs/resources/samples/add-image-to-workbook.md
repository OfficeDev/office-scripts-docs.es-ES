---
title: Agregar imágenes a un libro
description: Obtenga información sobre cómo usar Office scripts para agregar una imagen a un libro y copiarla entre hojas.
ms.date: 07/12/2021
localization_priority: Normal
ms.openlocfilehash: 993444aa328356f872db90d1b9d2403bf28be4de
ms.sourcegitcommit: a86b91c7e104bb7c26efd56de53b9e3976a34828
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 07/12/2021
ms.locfileid: "53394419"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="d3061-103">Agregar imágenes a un libro</span><span class="sxs-lookup"><span data-stu-id="d3061-103">Add images to a workbook</span></span>

<span data-ttu-id="d3061-104">En este ejemplo se muestra cómo trabajar con imágenes mediante un script Office en Excel.</span><span class="sxs-lookup"><span data-stu-id="d3061-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="d3061-105">Escenario</span><span class="sxs-lookup"><span data-stu-id="d3061-105">Scenario</span></span>

<span data-ttu-id="d3061-106">Las imágenes ayudan con la personalción de marca, la identidad visual y las plantillas.</span><span class="sxs-lookup"><span data-stu-id="d3061-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="d3061-107">Ayudan a hacer que un libro sea más que una tabla gigante.</span><span class="sxs-lookup"><span data-stu-id="d3061-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="d3061-108">El primer ejemplo copia una imagen de una hoja de cálculo a otra.</span><span class="sxs-lookup"><span data-stu-id="d3061-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="d3061-109">Esto podría usarse para colocar el logotipo de su empresa en la misma posición en cada hoja.</span><span class="sxs-lookup"><span data-stu-id="d3061-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="d3061-110">El segundo ejemplo copia una imagen de una dirección URL.</span><span class="sxs-lookup"><span data-stu-id="d3061-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="d3061-111">Esto podría usarse para copiar las fotos que un compañero almacenaba en una carpeta compartida en un libro relacionado.</span><span class="sxs-lookup"><span data-stu-id="d3061-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="d3061-112">Archivo Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="d3061-112">Sample Excel file</span></span>

<span data-ttu-id="d3061-113">Descargue <a href="add-images.xlsx">add-images.xlsx</a> para un libro listo para usar.</span><span class="sxs-lookup"><span data-stu-id="d3061-113">Download <a href="add-images.xlsx">add-images.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="d3061-114">Agregue los siguientes scripts y pruebe el ejemplo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="d3061-114">Add the following scripts and try the sample yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="d3061-115">Código de ejemplo: copiar una imagen en hojas de cálculo</span><span class="sxs-lookup"><span data-stu-id="d3061-115">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="d3061-116">Código de ejemplo: Agregar una imagen de una dirección URL a un libro</span><span class="sxs-lookup"><span data-stu-id="d3061-116">Sample code: Add an image from a URL to a workbook</span></span>

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
