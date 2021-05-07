---
title: Usar archivos de macro en Power Automate flujos
description: Obtenga información sobre cómo usar archivos de macros o archivos xlsm en Power Automate flujos.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b232a1d31a7ff6e28016c5e28fd8a83c8d3f1859
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232658"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="40a12-103">Cómo usar archivos de macro en Power Automate flujos</span><span class="sxs-lookup"><span data-stu-id="40a12-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="40a12-104">[Power Automate proporcionan](https://flow.microsoft.com/) [conectores](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) Excel para ayudar Excel conectar archivos Excel con el resto de los datos y aplicaciones de la organización, como Teams, Outlook y SharePoint.</span><span class="sxs-lookup"><span data-stu-id="40a12-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="40a12-105">Sin embargo, los archivos de macro no se pueden seleccionar en el desplegable de archivos (vea un ejemplo en la siguiente captura de pantalla).</span><span class="sxs-lookup"><span data-stu-id="40a12-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="La Power Automate ejecutar script que no muestra ningún archivo de macro seleccionado. El error que se muestra es &quot;Archivo&quot; es obligatorio":::

<span data-ttu-id="40a12-107">Una forma de evitar este problema es incluir la acción "Obtener metadatos de archivo" (OneDrive o SharePoint) y usar la propiedad ID en la acción "Ejecutar script", como se muestra en la siguiente captura de pantalla.</span><span class="sxs-lookup"><span data-stu-id="40a12-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="La Power Automate ejecutar script que muestra el archivo de macro seleccionado y no se produce ningún error de script de ejecución":::

> [!NOTE]
> <span data-ttu-id="40a12-109">Es posible que algunos XLSM (especialmente los que tienen controles ActiveX/Formulario) no funcionen en el Excel en línea.</span><span class="sxs-lookup"><span data-stu-id="40a12-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="40a12-110">Asegúrese de probar antes de implementar la solución.</span><span class="sxs-lookup"><span data-stu-id="40a12-110">Be sure to test before deploying your solution.</span></span>

## <a name="other-resources"></a><span data-ttu-id="40a12-111">Otros recursos</span><span class="sxs-lookup"><span data-stu-id="40a12-111">Other resources</span></span>

<span data-ttu-id="40a12-112">[Vea el vídeo de YouTube de Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)sobre cómo usar un archivo .xlsm en una acción Ejecutar script .</span><span class="sxs-lookup"><span data-stu-id="40a12-112">[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).</span></span>
