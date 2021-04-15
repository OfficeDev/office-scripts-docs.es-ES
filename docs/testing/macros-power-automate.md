---
title: Usar archivos de macro en flujos de Power Automate
description: Obtenga información sobre cómo usar archivos de macro o archivos xlsm en flujos de Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: a7929fc485ae2118d30a4f2783538d0e04deca2a
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755017"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Cómo usar archivos de macro en flujos de Power Automate

[Los flujos de Power Automate](https://flow.microsoft.com/) proporcionan [conectores de Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) para ayudar a conectar archivos de Excel con el resto de los datos y aplicaciones de la organización, como Teams, Outlook y SharePoint.

Sin embargo, los archivos de macro no se pueden seleccionar en el desplegable de archivos (vea un ejemplo en la siguiente captura de pantalla).

:::image type="content" source="../images/no-xlsm.png" alt-text="La acción de script Ejecutar de Power Automate que no muestra ningún archivo de macro seleccionado. El error que se muestra es &quot;Archivo&quot; es obligatorio.":::

Una forma de evitar este problema es incluir la acción "Obtener metadatos de archivo" (OneDrive o SharePoint) y usar la propiedad ID en la acción "Ejecutar script", como se muestra en la siguiente captura de pantalla.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="La acción de script Ejecutar de Power Automate que muestra el archivo de macro seleccionado y ningún error ejecutar script.":::

> [!NOTE]
> Es posible que algunos XLSM (especialmente los que tienen controles ActiveX/Formulario) no funcionen en el conector en línea de Excel. Asegúrese de probar antes de implementar la solución.

[![Ver vídeo sobre el uso de XLSM en la acción Ejecutar script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vídeo sobre el uso de XLSM en la acción Ejecutar script")
