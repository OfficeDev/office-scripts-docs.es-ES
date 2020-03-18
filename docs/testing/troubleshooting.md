---
title: Solución de problemas de scripts de Office
description: Sugerencias y técnicas de depuración de scripts de Office, así como recursos de ayuda.
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700356"
---
# <a name="troubleshooting-office-scripts"></a>Solución de problemas de scripts de Office

Al desarrollar scripts de Office, puede cometer errores. Es correcto. Tenemos herramientas que ayudan a encontrar los problemas y que los scripts funcionan perfectamente.

## <a name="console-logs"></a>Registros de la consola

En ocasiones, durante la solución de problemas, querrá imprimir los mensajes en la pantalla. Estos pueden mostrar el valor actual de las variables o las rutas de código que se están desencadenando. Para ello, registre el texto en la consola.

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> No olvide los `load` datos de la `sync` hoja de cálculo y con el libro antes de registrar las propiedades del objeto.

Las cadenas pasadas`console.log` a se mostrarán en la consola de registro del editor de código. Para activar la consola, presione el botón de **puntos suspensivos** y seleccione **registros...**

Los registros no afectan al libro.

## <a name="error-messages"></a>Mensajes de error

Cuando el script de Excel encuentra un problema en ejecución, produce un error. Verá un mensaje emergente en el que se le preguntará si desea **ver los registros**. Presione ese botón para abrir la consola y mostrar los errores.

## <a name="help-resources"></a>Recursos de ayuda

[Desbordamiento de pila](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores que desea ayudar con los problemas de codificación. A menudo, podrá encontrar la solución a su problema mediante una búsqueda rápida de desbordamiento de pila. Si no es así, formule su pregunta y etiquete con la etiqueta "Office-scripts". No olvide mencionar que está creando un *script*de Office, no un *complemento de*Office.

Si encuentra un problema con la API de JavaScript de Office, cree un problema en el repositorio de github [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) . Los miembros del equipo de producto responderán a los problemas y proporcionarán asistencia. La creación de un problema en el repositorio de **OfficeDev/Office-js** indica que ha encontrado un error en la biblioteca de la API de JavaScript de Office que el equipo del producto debe tratar.

Si hay un problema con el grabador de acciones o con el editor, envíe sus comentarios a través del botón **ayuda > comentarios** de Excel.

## <a name="see-also"></a>Vea también

- [Scripts de Office en Excel en la web](../overview/excel.md)
- [Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web](../develop/scripting-fundamentals.md)
- [Deshacer los efectos de un script de Office](undo.md)
