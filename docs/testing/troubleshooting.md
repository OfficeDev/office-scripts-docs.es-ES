---
title: Solucionar problemas Office scripts
description: Sugerencias y técnicas de depuración para Office scripts, así como recursos de ayuda.
ms.date: 11/11/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2c43d0e4b1f4fd5675397fd79eaab1345ae39b98
ms.sourcegitcommit: 24a6b8ae0cc57a0307fbc9b3e87432f1f4a92263
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 11/17/2021
ms.locfileid: "61064205"
---
# <a name="troubleshoot-office-scripts"></a>Solucionar problemas Office scripts

A medida que desarrolla Office scripts, puede cometer errores. Está bien. Tiene las herramientas para ayudar a encontrar los problemas y hacer que los scripts funcionen perfectamente.

> [!NOTE]
> Para obtener consejos de solución de problemas específicos Office scripts con Power Automate, vea [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).

## <a name="types-of-errors"></a>Tipos de errores

Office errores de scripts se en una de dos categorías:

* Errores o advertencias en tiempo de compilación
* Errores en tiempo de ejecución

### <a name="compile-time-errors"></a>Errores en tiempo de compilación

Los errores y advertencias en tiempo de compilación se muestran inicialmente en el Editor de código. Estos se muestran con los subrayados rojos ondulados del editor. También se muestran en la pestaña **Problemas** en la parte inferior del panel de tareas Editor de código. Al seleccionar el error, se darán más detalles sobre el problema y se sugerirán soluciones. Los errores en tiempo de compilación deben solucionarse antes de ejecutar el script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Error del compilador que se muestra en el texto activado del Editor de código.":::

También puede ver subrayados de advertencia naranja y mensajes informativos grises. Estas indican sugerencias de rendimiento u otras posibilidades en las que el script puede tener efectos involuntarias. Estas advertencias deben examinarse estrechamente antes de descartarlas.

### <a name="runtime-errors"></a>Errores en tiempo de ejecución

Los errores en tiempo de ejecución se producen debido a problemas de lógica en el script. Esto podría deberse a que un objeto usado en el script no está en el libro, una tabla tiene un formato diferente al previsto o alguna otra discrepancia leve entre los requisitos del script y el libro actual. El siguiente script genera un error cuando no está presente una hoja de cálculo denominada "TestSheet".

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Mensajes de consola

Tanto los errores en tiempo de compilación como en tiempo de ejecución muestran mensajes de error en la consola cuando se ejecuta un script. Dan un número de línea donde se encontró el problema. Tenga en cuenta que la causa raíz de cualquier problema puede ser una línea de código diferente a la que se indica en la consola.

En la imagen siguiente se muestra el resultado de la consola del [error explícito `any` ](../develop/typescript-restrictions.md) del compilador. Tenga en cuenta `[5, 16]` el texto al principio de la cadena de error. Esto indica que el error está en la línea 5, empezando por el carácter 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="La consola del Editor de código que muestra un mensaje de error explícito &quot;cualquiera&quot;.":::

La imagen siguiente muestra el resultado de la consola de un error en tiempo de ejecución. Aquí, el script intenta agregar una hoja de cálculo con el nombre de una hoja de cálculo existente. De nuevo, anote la "Línea 2" anterior al error para mostrar la línea que se debe investigar.
:::image type="content" source="../images/runtime-error-console.png" alt-text="La consola del Editor de código que muestra un error de la llamada &quot;addWorksheet&quot;.":::

## <a name="console-logs"></a>Registros de consola

Imprimir mensajes en la pantalla con la `console.log` instrucción. Estos registros pueden mostrar el valor actual de las variables o qué rutas de código se están desencadenando. Para ello, llame `console.log` con cualquier objeto como parámetro. Por lo general, a `string` es el tipo más fácil de leer en la consola.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Las cadenas pasadas se muestran en la consola de registro del Editor de `console.log` código, en la parte inferior del panel de tareas. Los registros se encuentran en la **pestaña Salida,** aunque la pestaña aumenta automáticamente el foco cuando se escribe un registro.

Los registros no afectan al libro.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>La pestaña Automatizar no aparece ni Office scripts no están disponibles

Los siguientes pasos deben ayudar a solucionar los problemas relacionados con la pestaña **Automatizar** que no aparezcan en Excel en la Web.

1. [Asegúrese de que su Microsoft 365 incluye Office scripts](../overview/excel.md#requirements).
1. [Compruebe que el explorador es compatible.](platform-limits.md#browser-support)
1. [Asegúrese de que las cookies de terceros están habilitadas](platform-limits.md#third-party-cookies).
1. [Asegúrese de que el administrador no ha deshabilitado Office scripts en el Centro de administración de Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a>Recursos de ayuda

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores dispuestos a ayudar con problemas de codificación. A menudo, podrás encontrar la solución al problema mediante una búsqueda rápida de desbordamiento de pila. Si no es así, haga su pregunta y etiquete con la etiqueta "office-scripts". Asegúrese de mencionar que está creando un script Office *,* no un Office *complemento*.

## <a name="see-also"></a>Consulte también

- [Procedimientos recomendados para Scripts de Office](../develop/best-practices.md)
- [Límites de plataforma con Office scripts](platform-limits.md)
- [Mejorar el rendimiento de los scripts Office scripts](../develop/web-client-performance.md)
- [Solucionar Office scripts que se ejecutan en PowerAutomate](power-automate-troubleshooting.md)
- [Deshacer los efectos de Scripts de Office](undo.md)
