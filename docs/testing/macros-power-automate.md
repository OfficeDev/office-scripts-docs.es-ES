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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Cómo usar archivos de macro en flujos de Power Automate

[Los flujos de Power Automate](https://flow.microsoft.com/) proporcionan [conectores de Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) para ayudar a conectar archivos de Excel con el resto de los datos y aplicaciones de la organización, como Teams, Outlook y SharePoint.

Sin embargo, los archivos de macro no se pueden seleccionar en el desplegable de archivos (vea un ejemplo en la siguiente captura de pantalla).

![Sin xlsm en la acción Ejecutar script](../images/no-xlsm.png)

Una forma de evitar este problema es incluir la acción "Obtener metadatos de archivo" (OneDrive o SharePoint) y usar la propiedad ID en la acción "Ejecutar script", como se muestra en la siguiente captura de pantalla.

![xlsm en la acción Ejecutar script](../images/xlsm-in-pa.png)

> [!NOTE]
> Es posible que algunos XLSM (especialmente los que tienen controles ActiveX/Formulario) no funcionen en el conector en línea de Excel. Asegúrese de probar antes de implementar la solución.

[![Ver vídeo sobre el uso de XLSM en la acción Ejecutar script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vídeo sobre el uso de XLSM en la acción Ejecutar script")
