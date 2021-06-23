---
title: Usar archivos de macro en Power Automate flujos
description: Obtenga información sobre cómo usar archivos de macros o archivos xlsm en Power Automate flujos.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 91e11424e4220a3e1f80cdd2711d05f219016147
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074644"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Cómo usar archivos de macro en Power Automate flujos

[Power Automate proporcionan](https://flow.microsoft.com/) [conectores](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) Excel para ayudar Excel conectar archivos Excel con el resto de los datos y aplicaciones de la organización, como Teams, Outlook y SharePoint.

Sin embargo, los archivos de macro no se pueden seleccionar en el desplegable de archivos (vea un ejemplo en la siguiente captura de pantalla).

:::image type="content" source="../images/no-xlsm.png" alt-text="La Power Automate ejecutar script que no muestra ningún archivo de macro seleccionado. El error que se muestra es &quot;Archivo&quot; es obligatorio.":::

Una forma de evitar este problema es incluir la acción "Obtener metadatos de archivo" (OneDrive o SharePoint) y usar la propiedad ID en la acción "Ejecutar script", como se muestra en la siguiente captura de pantalla.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="La Power Automate ejecutar script que muestra el archivo de macro seleccionado y ningún error ejecutar script.":::

> [!NOTE]
> Es posible que algunos XLSM (especialmente los que tienen controles ActiveX/Formulario) no funcionen en el Excel en línea. Asegúrese de probar antes de implementar la solución.

## <a name="other-resources"></a>Otros recursos

[Vea el vídeo de YouTube de Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)sobre cómo usar un archivo .xlsm en una acción Ejecutar script .
