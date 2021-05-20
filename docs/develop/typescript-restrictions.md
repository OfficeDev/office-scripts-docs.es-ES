---
title: Restricciones de TypeScript en scripts de Office
description: Los detalles del compilador typescript y linter utilizados por el Editor de código de scripts de Office.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: a4198e0e56224ac5da89e89c43c8d2f3ef44d6d7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545022"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Restricciones de TypeScript en scripts de Office

Office Los scripts utilizan el lenguaje TypeScript. En su mayor parte, cualquier código TypeScript o JavaScript funcionará en scripts Office. Sin embargo, el Editor de código aplica algunas restricciones para garantizar que el script funcione de forma coherente y según lo previsto con el libro de trabajo Excel.

## <a name="no-any-type-in-office-scripts"></a>No hay ningún tipo de script de Office

Escribir [tipos](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) es opcional en TypeScript, porque los tipos se pueden inferir. Sin embargo, Office Scripts requiere que una variable no pueda ser de [tipo .](https://www.typescriptlang.org/docs/handbook/basic-types.html#any) Tanto explícito como implícito `any` no se permiten en Office scripts. Estos casos se notifican como errores.

### <a name="explicit-any"></a>explícito `any`

No puede declarar explícitamente que una variable sea de tipo `any` en scripts de Office (es decir, `let someVariable: any;` ). El `any` tipo causa problemas cuando se procesa por Excel. Por ejemplo, `Range` es necesario saber que un valor es un , o `string` `number` `boolean` . Recibirá un error en tiempo de compilación (un error antes de ejecutar el script) si alguna variable se define explícitamente como el `any` tipo del script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="El mensaje explícito de &quot;cualquiera&quot; en el texto flotante del Editor de código":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="El error explícito de &quot;cualquiera&quot; en la ventana de la consola":::

En la captura de pantalla anterior `[5, 16] Explicit Any is not allowed` indica que la línea #5, la columna #16 define el `any` tipo. Esto le ayuda a localizar el error.

Para evitar este problema, defina siempre el tipo de la variable. Si no está seguro sobre el tipo de variable, puede utilizar un [tipo de unión](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html). Esto puede ser útil para las variables que mantienen `Range` valores, que pueden ser de tipo `string` , `number` o `boolean` (el tipo de valores es una unión `Range` de esos: `string | number | boolean` ).

### <a name="implicit-any"></a>implícito `any`

Los tipos de variable TypeScript se pueden definir [implícitamente.](https://www.typescriptlang.org/docs/handbook/type-inference.html) Si el compilador TypeScript no puede determinar el tipo de una variable (ya sea porque el tipo no se define explícitamente o la inferencia de tipo no es posible), entonces es un error implícito `any` y recibirá un error en tiempo de compilación.

El caso más común en cualquier implícito `any` está en una declaración de variable, como `let value;` . Hay dos maneras de evitar esto:

* Asigne la variable a un tipo implícitamente identificable ( `let value = 5;` o `let value = workbook.getWorksheet();` ).
* Escriba explícitamente la variable ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Sin heredar clases o interfaces de script Office

Las clases e interfaces que se crean en el script de Office no pueden [extender ni implementar](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office clases o interfaces de scripts. En otras palabras, nada en el `ExcelScript` espacio de nombres puede tener subclases o subinterfaces.

## <a name="incompatible-typescript-functions"></a>Funciones incompatibles typescript

Office Las API de scripts no se pueden utilizar en lo siguiente:

* [Funciones del generador](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` no se admite

La [función eval de](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript no es compatible por razones de seguridad.

## <a name="restricted-identifers"></a>Identificadores restringidos

Las siguientes palabras no se pueden usar como identificadores en un script. Son términos reservados.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Solo funciones de flecha en devoluciones de llamada de matriz

Los scripts solo pueden usar [funciones de flecha](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) al proporcionar argumentos de devolución de llamada para los métodos [Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) No puede pasar ningún tipo de identificador o función "tradicional" a estos métodos.

```TypeScript
const myArray = [1, 2, 3, 4, 5, 6];
let filteredArray = myArray.filter((x) => {
  return x % 2 === 0;
});
/*
  The following code generates a compiler error in the Office Scripts Code Editor.
  filteredArray = myArray.filter(function (x) {
    return x % 2 === 0;
  });
*/
```

## <a name="performance-warnings"></a>Advertencias de rendimiento

Linter del [](https://wikipedia.org/wiki/Lint_(software)) Editor de código proporciona advertencias si el script podría tener problemas de rendimiento. Los casos y cómo solucionarlos se documentan en [Mejorar el rendimiento de los scripts de Office.](web-client-performance.md)

## <a name="external-api-calls"></a>Llamadas a la API externa

Consulte [Compatibilidad con llamadas a la API externa en scripts de Office](external-calls.md) para obtener más información.

## <a name="see-also"></a>Vea también

* [Conceptos básicos de los Scripts de Office en Excel en la web](scripting-fundamentals.md)
* [Mejore el rendimiento de sus scripts de Office](web-client-performance.md)
