---
title: Solución de problemas de scripts Office
description: Sugerencias y técnicas de depuración para scripts de Office, así como recursos de ayuda.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545562"
---
# <a name="troubleshoot-office-scripts"></a>Solución de problemas de scripts Office

A medida que desarrolla Office scripts, puede cometer errores. Está bien. Usted tiene las herramientas para ayudar a encontrar los problemas y hacer que sus scripts funcionen perfectamente.

## <a name="types-of-errors"></a>Tipos de errores

Office Los errores de scripts se dividen en una de las dos categorías:

* Compilar errores o advertencias en tiempo de compilación
* Errores en tiempo de ejecución

### <a name="compile-time-errors"></a>Errores de tiempo de compilación

Los errores y advertencias en tiempo de compilación se muestran inicialmente en el Editor de código. Estos se muestran por los subrayados rojos ondulados en el editor. También se muestran en la pestaña **Problemas** en la parte inferior del panel de tareas Editor de código. La selección del error dará más detalles sobre el problema y sugerirá soluciones. Los errores en tiempo de compilación deben solucionarse antes de ejecutar el script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Error del compilador que se muestra en el texto flotante del Editor de código":::

También puede ver subrayados de advertencia naranja y mensajes informativos grises. Estos indican sugerencias de rendimiento u otras posibilidades donde el script puede tener efectos involuntarios. Estas advertencias deben examinarse detenidamente antes de desestimarlas.

### <a name="runtime-errors"></a>Errores en tiempo de ejecución

Los errores en tiempo de ejecución se producen debido a problemas lógicos en el script. Esto podría deberse a que un objeto utilizado en el script no está en el libro, una tabla tiene un formato diferente al previsto o alguna otra discrepancia leve entre los requisitos del script y el libro de trabajo actual. El siguiente script genera un error cuando una hoja de cálculo denominada "TestSheet" no está presente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Mensajes de consola

Los errores en tiempo de compilación y tiempo de ejecución muestran mensajes de error en la consola cuando se ejecuta un script. Dan un número de línea donde se encontró el problema. Tenga en cuenta que la causa raíz de cualquier problema puede ser una línea de código diferente a la indicada en la consola.

La siguiente imagen muestra la salida de la consola para el error [explícito `any` ](../develop/typescript-restrictions.md) del compilador. Anote el texto `[5, 16]` al principio de la cadena de error. Esto indica que el error está en la línea 5, comenzando en el carácter 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="La consola del Editor de código que muestra un mensaje de error explícito de &quot;cualquier&quot;":::

La siguiente imagen muestra la salida de la consola para un error en tiempo de ejecución. Aquí, el script intenta agregar una hoja de cálculo con un nombre de una hoja de cálculo existente. Una vez más, observe la "Línea 2" anterior al error para mostrar qué línea investigar.
:::image type="content" source="../images/runtime-error-console.png" alt-text="La consola del Editor de código que muestra un error de la llamada a 'addWorksheet'":::

## <a name="console-logs"></a>Registros de consola

Imprima mensajes en la pantalla con la `console.log` instrucción. Estos registros pueden mostrar el valor actual de las variables o qué rutas de código se están desencadenando. Para ello, llame `console.log` con cualquier objeto como parámetro. Por lo general, a `string` es el tipo más fácil de leer en la consola.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Las cadenas a las que `console.log` se pasa se muestran en la consola de registro del Editor de código, en la parte inferior del panel de tareas. Los registros se encuentran en la pestaña **Salida,** aunque la pestaña gana automáticamente el foco cuando se escribe un registro.

Los registros no afectan al libro.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Automatice la pestaña que no aparece o Office scripts no disponibles

Los pasos siguientes deben ayudar a solucionar cualquier problema relacionado con la pestaña **Automatizar** que no aparezca en Excel en la Web.

1. [Asegúrese de que la licencia de Microsoft 365 incluya scripts Office](../overview/excel.md#requirements).
1. [Compruebe que su navegador es compatible.](platform-limits.md#browser-support)
1. [Asegúrese de que las cookies de terceros estén habilitadas.](platform-limits.md#third-party-cookies)
1. [Asegúrese de que el administrador no ha deshabilitado Office scripts en el Centro de administración de Microsoft 365.](/microsoft-365/admin/manage/manage-office-scripts-settings)

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>Solucionar problemas de scripts en Power Automate

Para obtener información específica para ejecutar scripts a través de Power Automate, consulte [Solución de problemas Office scripts que se ejecutan en Power Automate](power-automate-troubleshooting.md).

## <a name="help-resources"></a>Ayudar a los recursos

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores dispuestos a ayudar con los problemas de codificación. A menudo, podrás encontrar la solución a tu problema a través de una búsqueda rápida de Stack Overflow. Si no es así, haga su pregunta y etiquete con la etiqueta "office-scripts". Asegúrese de mencionar que está creando un *script* Office, no un *complemento Office*.

Si tiene un problema con la API de JavaScript Office, cree un problema en el repositorio de [officedev/office-js](https://github.com/OfficeDev/office-js) GitHub. Los miembros del equipo del producto responderán a las cuestiones y proporcionarán más asistencia. La creación de un problema en el repositorio **OfficeDev/office-js** indica que ha encontrado un defecto en la biblioteca de API de JavaScript Office que el equipo del producto debe abordar.

Si hay un problema con el Grabador de acciones o editor, envíe comentarios a través del botón **Ayuda > Comentarios** en Excel.

## <a name="see-also"></a>Vea también

- [Procedimientos recomendados para Scripts de Office](../develop/best-practices.md)
- [Límites de plataforma con scripts de Office](platform-limits.md)
- [Mejore el rendimiento de sus scripts de Office](../develop/web-client-performance.md)
- [Solucionar problemas Office scripts que se ejecutan en PowerAutomate](power-automate-troubleshooting.md)
- [Deshacer los efectos de Scripts de Office](undo.md)
