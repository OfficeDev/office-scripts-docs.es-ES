---
title: Usar archivos de macro en Power Automate flujos
description: Obtenga información sobre cómo usar archivos de macros o archivos xlsm en Power Automate flujos.
ms.date: 09/01/2021
localization_priority: Normal
ms.openlocfilehash: 204eb8315f90c0ab0e20a34ec64517fbfbf304b1
ms.sourcegitcommit: 9472e78eb186ceffdaaceb2718d5074ce55a0e74
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2021
ms.locfileid: "58866547"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Cómo usar archivos de macro en Power Automate flujos

El [Excel online (empresa)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) de [Power Automate](https://flow.microsoft.com/) normalmente solo funciona con archivos en el formato de hoja de cálculo Microsoft Excel Open XML (.xlsx). El explorador de archivos limita la selección .xlsx archivos dentro del conector. Sin embargo, los archivos de macro son compatibles con la acción **ejecutar script** del conector si se usan los metadatos del archivo.

En el flujo, use la acción **Obtener** metadatos de archivo desde [los conectores OneDrive para la Empresa](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) o [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) archivo. La **acción Ejecutar script** acepta estos metadatos como un archivo válido. Use el *contenido dinámico id.* devuelto desde la **acción Obtener metadatos de** archivo como argumento "Archivo" al ejecutar el script. La siguiente captura de pantalla muestra un flujo que proporciona los metadatos de un archivo denominado "Test Macro File.xlsm" a una **acción ejecutar script.**

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Flujo con una acción Obtener metadatos de archivo que pasa los metadatos de un archivo de macro a una acción ejecutar script.":::

> [!WARNING]
> Es posible que algunos archivos .xlsm, especialmente aquellos con controles ActiveX o Formulario, no funcionen en el Excel en línea. Asegúrese de probar antes de implementar la solución.

## <a name="other-resources"></a>Otros recursos

[Vea el vídeo de YouTube de Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)sobre cómo usar un archivo .xlsm en una acción Ejecutar script .
