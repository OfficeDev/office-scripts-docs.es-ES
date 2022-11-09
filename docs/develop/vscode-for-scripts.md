---
title: Visual Studio Code para scripts de Office (versión preliminar)
description: Cómo configurar el Editor de código de scripts de Office para conectarse a VS Code for the Web.
ms.date: 11/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: fd9dd417610c8ad64fbd3fc50048ce56afdb4e28
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 11/09/2022
ms.locfileid: "68892056"
---
# <a name="visual-studio-code-for-office-scripts-preview"></a>Visual Studio Code para scripts de Office (versión preliminar)

[Visual Studio Code para web](https://vscode.dev/) permite a los usuarios editar cualquier cosa desde cualquier lugar. Conecte la experiencia de Scripts de Office a este editor de código popular para iniciar el scripting fuera del libro.

:::image type="content" source="../images/vscode-script-editor.png" alt-text="Una ventana Excel en la Web con el Editor de código abierto junto a un vs code en la ventana web con un script abierto.":::

Visual Studio Code tiene algunas ventajas sobre el Editor de código integrado.

- Edición a pantalla completa! El script ya no tiene que compartir espacio de pantalla con el libro.
- Edite varios scripts a la vez. Cambie rápidamente entre scripts para compartir código de otras automatizaciones.
- ¡Extensiones! Use sus extensiones de VS Code favoritas para la revisión ortográfica, el formato y cualquier otra cosa que le ayude a realizar el trabajo.

> [!NOTE]
> Esta característica está en versión preliminar. Está sujeto a cambios en función de los comentarios. Si encuentra algún problema, informe de ellos a través del botón **Comentarios** en Excel. A continuación se muestran problemas conocidos con la versión actual de la característica.
>
> - Visual Studio Code solo se puede conectar a scripts de Office a través de Excel en la Web.
> - Esta conexión de Scripts de Office solo está disponible con clientes de Excel en inglés.

## <a name="connect-visual-studio-code-to-office-scripts"></a>Conexión de Visual Studio Code a scripts de Office

Siga estos pasos únicos para conectar Visual Studio Code y Excel en la Web.

1. Abra el Editor de **código** de scripts de Office.
2. En el menú **Más opciones (...) ,** seleccione **Configuración del editor**.
3. Seleccione **(versión preliminar) Visual Studio Code conexión**.

:::image type="content" source="../images/vscode-enable-option.png" alt-text="Panel de tareas de configuración del editor que muestra una casilla etiquetada Visual Studio Code conexión.":::

Ahora puede editar y ejecutar los scripts desde Visual Studio Code. En cualquier script, vaya al menú **Más opciones (...)** y seleccione **Abrir en VS Code**.

:::image type="content" source="../images/vscode-open-option.png" alt-text="La opción Abrir en VS Code que se está seleccionando en una lista junto a un script abierto.":::

## <a name="see-also"></a>Vea también

- [Entorno del Editor de código de scripts de Office](../overview/code-editor-environment.md)
- [Visual Studio Code para la Web (documentación)](https://code.visualstudio.com/docs/editor/vscode-web)
