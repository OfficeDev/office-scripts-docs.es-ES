---
title: Solución de problemas de scripts de Office
description: Sugerencias y técnicas de depuración para scripts de Office, así como recursos de ayuda.
ms.date: 11/11/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e673d39b6249ccc7598b832d6478cc8dc0751f6
ms.sourcegitcommit: f5fc9146d5c096e3a580a3fa8f9714147c548df4
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/12/2022
ms.locfileid: "66038682"
---
# <a name="troubleshoot-office-scripts"></a>Solución de problemas de scripts de Office

A medida que desarrolle scripts Office, puede cometer errores. Está bien. Tiene las herramientas para ayudar a encontrar los problemas y conseguir que los scripts funcionen perfectamente.

> [!NOTE]
> Para obtener consejos de solución de problemas específicos de scripts de Office con Power Automate, consulte [Solución de problemas de scripts Office que se ejecutan en Power Automate](power-automate-troubleshooting.md).

## <a name="types-of-errors"></a>Tipos de errores

Office errores de scripts se dividen en una de dos categorías:

* Errores o advertencias en tiempo de compilación
* Errores en tiempo de ejecución

### <a name="compile-time-errors"></a>Errores en tiempo de compilación

Los errores y advertencias en tiempo de compilación se muestran inicialmente en el Editor de código. Estos se muestran mediante los subrayados rojos ondulados en el editor. También se muestran en la pestaña **Problemas** de la parte inferior del panel de tareas Editor de código. Al seleccionar el error, se proporcionarán más detalles sobre el problema y se sugerirán soluciones. Los errores en tiempo de compilación deben solucionarse antes de ejecutar el script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Error del compilador que se muestra en el texto del mouse del Editor de código.":::

También puede ver subrayados de advertencia naranja y mensajes informativos grises. Estos indican sugerencias de rendimiento u otras posibilidades en las que el script puede tener efectos involuntarios. Estas advertencias deben examinarse detenidamente antes de descartarlas.

### <a name="runtime-errors"></a>Errores en tiempo de ejecución

Los errores en tiempo de ejecución se producen debido a problemas lógicos en el script. Esto podría deberse a que un objeto usado en el script no está en el libro, una tabla tiene un formato diferente al previsto o a alguna otra ligera discrepancia entre los requisitos del script y el libro actual. El siguiente script genera un error cuando una hoja de cálculo denominada "TestSheet" no está presente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Mensajes de consola

Los errores en tiempo de compilación y en tiempo de ejecución muestran mensajes de error en la consola cuando se ejecuta un script. Proporcionan un número de línea donde se encontró el problema. Tenga en cuenta que la causa principal de cualquier problema puede ser una línea de código diferente de la indicada en la consola.

En la imagen siguiente se muestra la salida de la consola para el error [explícito `any`](../develop/typescript-restrictions.md) del compilador. Anote el texto `[5, 16]` al principio de la cadena de error. Esto indica que el error está en la línea 5, empezando por el carácter 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="La consola del Editor de código que muestra un mensaje de error explícito &quot;any&quot;.":::

En la imagen siguiente se muestra la salida de la consola de un error en tiempo de ejecución. Aquí, el script intenta agregar una hoja de cálculo con el nombre de una hoja de cálculo existente. Una vez más, observe la "línea 2" anterior al error para mostrar qué línea investigar.
:::image type="content" source="../images/runtime-error-console.png" alt-text="La consola del Editor de código muestra un error de la llamada &quot;addWorksheet&quot;.":::

## <a name="console-logs"></a>Registros de consola

Imprima mensajes en la pantalla con la `console.log` instrucción . Estos registros pueden mostrar el valor actual de las variables o qué rutas de acceso de código se desencadenan. Para ello, llame a `console.log` con cualquier objeto como parámetro. Normalmente, es `string` el tipo más fácil de leer en la consola.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Las cadenas pasadas a `console.log` se muestran en la consola de registro del Editor de código, en la parte inferior del panel de tareas. Los registros se encuentran en la pestaña **Salida** , aunque la pestaña obtiene automáticamente el foco cuando se escribe un registro.

Los registros no afectan al libro.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>La pestaña Automatizar no aparece ni Office scripts no disponibles

Los pasos siguientes deben ayudar a solucionar cualquier problema relacionado con la pestaña **Automatizar** que no aparezca en Excel en la Web.

1. [Asegúrese de que la licencia de Microsoft 365 incluye scripts de Office](../overview/excel.md#requirements).
1. [Compruebe que el explorador es compatible](platform-limits.md#browser-support).
1. [Asegúrese de que las cookies de terceros están habilitadas](platform-limits.md#third-party-cookies).
1. [Asegúrese de que el administrador no ha deshabilitado Office scripts en el Centro de administración de Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).
1. Asegúrese de que no ha iniciado sesión como usuario externo o invitado en el inquilino.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a>Recursos de ayuda

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores dispuestos a ayudar con los problemas de codificación. A menudo, podrá encontrar la solución al problema a través de una búsqueda rápida de Stack Overflow. Si no es así, formule su pregunta y etiquete con la etiqueta "office-scripts". Asegúrese de mencionar que está creando un *script* de Office, no un *complemento* de Office.

## <a name="see-also"></a>Vea también

- [Procedimientos recomendados para Scripts de Office](../develop/best-practices.md)
- [Límites de plataforma con scripts de Office](platform-limits.md)
- [Mejora del rendimiento de los scripts de Office](../develop/web-client-performance.md)
- [Solución de problemas de scripts de Office que se ejecutan en PowerAutomate](power-automate-troubleshooting.md)
- [Deshacer los efectos de Scripts de Office](undo.md)
