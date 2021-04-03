---
title: Usar archivos de macro en flujos de Power Automate
description: Obtenga información sobre cómo usar archivos de macro o archivos xlsm en flujos de Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: ec1fe00eb9ddc382ae4bc02187de7a36c97288b1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571610"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="ef1a1-103">Cómo usar archivos de macro en flujos de Power Automate</span><span class="sxs-lookup"><span data-stu-id="ef1a1-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="ef1a1-104">[Los flujos de Power Automate](https://flow.microsoft.com/) proporcionan [conectores de Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) para ayudar a conectar archivos de Excel con el resto de los datos y aplicaciones de la organización, como Teams, Outlook y SharePoint.</span><span class="sxs-lookup"><span data-stu-id="ef1a1-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="ef1a1-105">Sin embargo, los archivos de macro no se pueden seleccionar en el desplegable de archivos (vea un ejemplo en la siguiente captura de pantalla).</span><span class="sxs-lookup"><span data-stu-id="ef1a1-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

![Sin xlsm en la acción Ejecutar script](../images/no-xlsm.png)

<span data-ttu-id="ef1a1-107">Una forma de evitar este problema es incluir la acción "Obtener metadatos de archivo" (OneDrive o SharePoint) y usar la propiedad ID en la acción "Ejecutar script", como se muestra en la siguiente captura de pantalla.</span><span class="sxs-lookup"><span data-stu-id="ef1a1-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

![xlsm en la acción Ejecutar script](../images/xlsm-in-pa.png)

> [!NOTE]
> <span data-ttu-id="ef1a1-109">Es posible que algunos XLSM (especialmente los que tienen controles ActiveX/Formulario) no funcionen en el conector en línea de Excel.</span><span class="sxs-lookup"><span data-stu-id="ef1a1-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="ef1a1-110">Asegúrese de probar antes de implementar la solución.</span><span class="sxs-lookup"><span data-stu-id="ef1a1-110">Be sure to test before deploying your solution.</span></span>

<span data-ttu-id="ef1a1-111">[![Ver vídeo sobre el uso de XLSM en la acción Ejecutar script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vídeo sobre el uso de XLSM en la acción Ejecutar script")</span><span class="sxs-lookup"><span data-stu-id="ef1a1-111">[![Watch video about using XLSM in Run Script action](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video about using XLSM in Run Script action")</span></span>