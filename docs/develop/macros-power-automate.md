---
title: Usar archivos habilitados para macros en Power Automate flujos
description: Obtenga información sobre cómo usar archivos habilitados para macros o archivos .xlsm en Power Automate flujos.
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f2ecefe9fb97d1c5514ddb52c3cbcd0596df426
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585747"
---
# <a name="how-to-use-macro-enabled-files-in-power-automate-flows"></a>Cómo usar archivos habilitados para macros en Power Automate flujos

Puede integrar los archivos .xlsm con un flujo Power Automate datos. Esto le permite empezar a convertir las soluciones de automatización existentes a formatos basados en web. Tenga en cuenta que las macros contenidas en los archivos .xslm no se pueden ejecutar a través de Power Automate. Solo Office scripts están habilitados allí.

El [Excel online (empresa)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) de [Power Automate](https://flow.microsoft.com/) suele limitarse a los archivos en el formato Microsoft Excel hoja de cálculo de Open XML (.xlsx). Su explorador de archivos solo permite seleccionar .xlsx archivos. Sin embargo, los archivos habilitados para macros son compatibles con la acción **ejecutar script** del conector si se usan los metadatos del archivo.

En el flujo, use la acción Obtener metadatos **de** archivo desde los [conectores OneDrive para la Empresa](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) o [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) archivo. La **acción Ejecutar script** acepta estos metadatos como un archivo válido. Use el *contenido dinámico id* . devuelto desde la **acción Obtener metadatos de** archivo como argumento "Archivo" al ejecutar el script. La siguiente captura de pantalla muestra un flujo que proporciona los metadatos de un archivo denominado "Test Macro File.xlsm" a una **acción ejecutar script** .

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Flujo con una acción Obtener metadatos de archivo que pasa los metadatos de un archivo de macro a una acción ejecutar script.":::

> [!WARNING]
> Es posible que algunos archivos .xlsm, especialmente aquellos con controles ActiveX o Formulario, no funcionen en el Excel en línea. Asegúrese de probar antes de implementar la solución.

## <a name="other-resources"></a>Otros recursos

[Vea el vídeo de YouTube de Sudhi Ramamurthy sobre cómo usar un archivo .xlsm en una acción Ejecutar script](https://youtu.be/o-H9BbywJQQ).
