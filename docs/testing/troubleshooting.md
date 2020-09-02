---
title: Solución de problemas de scripts de Office
description: Sugerencias y técnicas de depuración de scripts de Office, así como recursos de ayuda.
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 0a2c9ab209bc36e8ba2bdb25a6ab79d9f900f29a
ms.sourcegitcommit: e9a8ef5f56177ea9a3d2fc5ac636368e5bdae1f4
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/01/2020
ms.locfileid: "47321596"
---
# <a name="troubleshooting-office-scripts"></a>Solución de problemas de scripts de Office

Al desarrollar scripts de Office, puede cometer errores. Es correcto. Tenemos herramientas que ayudan a encontrar los problemas y que los scripts funcionan perfectamente.

## <a name="console-logs"></a>Registros de la consola

En ocasiones, durante la solución de problemas, querrá imprimir los mensajes en la pantalla. Estos pueden mostrar el valor actual de las variables o las rutas de código que se están desencadenando. Para ello, registre el texto en la consola.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Las cadenas pasadas a `console.log` se mostrarán en la consola de registro del editor de código. Para activar la consola, presione el botón de **puntos suspensivos** y seleccione **registros...**

Los registros no afectan al libro.

## <a name="error-messages"></a>Mensajes de error

Cuando el script de Excel encuentra un problema en ejecución, produce un error. Verá un mensaje emergente en el que se le preguntará si desea **ver los registros**. Presione ese botón para abrir la consola y mostrar los errores.

## <a name="automate-tab-not-appearing"></a>La ficha automatizada no aparece

Los pasos siguientes le ayudarán a solucionar los problemas relacionados con la ficha **automatizar** que no aparecen en Excel en la Web.

1. Asegúrese [de que su licencia de 365 de Microsoft incluye scripts de Office](../overview/excel.md#requirements).
1. [Pida al administrador que habilite la característica](/microsoft-365/admin/manage/manage-office-scripts-settings).
1. [Compruebe que el explorador es compatible](platform-limits.md#browser-support).
1. [Asegúrese de que las cookies de terceros están habilitadas](platform-limits.md#third-party-cookies).

## <a name="help-resources"></a>Recursos de ayuda

[Desbordamiento de pila](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores que desea ayudar con los problemas de codificación. A menudo, podrá encontrar la solución a su problema mediante una búsqueda rápida de desbordamiento de pila. Si no es así, formule su pregunta y etiquete con la etiqueta "Office-scripts". No olvide mencionar que está creando un *script*de Office, no un *complemento de*Office.

Si encuentra un problema con la API de JavaScript de Office, cree un problema en el repositorio de github [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) . Los miembros del equipo de producto responderán a los problemas y proporcionarán asistencia. La creación de un problema en el repositorio de **OfficeDev/Office-js** indica que ha encontrado un error en la biblioteca de la API de JavaScript de Office que el equipo del producto debe tratar.

Si hay un problema con el grabador de acciones o con el editor, envíe sus comentarios a través del botón **ayuda > comentarios** de Excel.

## <a name="see-also"></a>Recursos adicionales

- [Scripts de Office en Excel en la Web](../overview/excel.md)
- [Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web](../develop/scripting-fundamentals.md)
- [Límites de plataforma con scripts de Office](platform-limits.md)
- [Mejorar el rendimiento de los scripts de Office](../develop/web-client-performance.md)
- [Deshacer los efectos de un script de Office](undo.md)
