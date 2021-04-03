---
title: Restricciones de TypeScript en scripts de Office
description: Los detalles del compilador TypeScript y linter usados por el Editor de código de scripts de Office.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 8c9d1beafb236e7ba10dedf00fab944c40fb954d
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570279"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Restricciones de TypeScript en scripts de Office

Los scripts de Office usan el lenguaje TypeScript. En su mayoría, cualquier código TypeScript o JavaScript funcionará en un script de Office. Sin embargo, el Editor de código aplica algunas restricciones para garantizar que el script funciona de forma coherente y según lo previsto con el libro de Excel.

## <a name="no-any-type-in-office-scripts"></a>No hay tipo 'any' en scripts de Office

Los [tipos](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) de escritura son opcionales en TypeScript, ya que los tipos se pueden deducir. Sin embargo, el script de Office requiere que una variable no pueda ser de [tipo .](https://www.typescriptlang.org/docs/handbook/basic-types.html#any) No se permiten explícitas `any` ni implícitas en un script de Office. Estos casos se notifican como errores.

### <a name="explicit-any"></a>Explícito `any`

No puede declarar explícitamente una variable como de tipo en scripts de `any` Office (es decir, `let someVariable: any;` ). El `any` tipo provoca problemas al procesarlo Excel. Por ejemplo, es `Range` necesario saber que un valor es un , o `string` `number` `boolean` . Recibirá un error en tiempo de compilación (un error antes de ejecutar el script) si alguna variable se define explícitamente como el tipo `any` en el script.

![El mensaje explícito de cualquier mensaje en el texto activado del editor de código](../images/explicit-any-editor-message.png)

![El error explícito de la ventana de la consola](../images/explicit-any-error-message.png)

En la captura de pantalla `[5, 16] Explicit Any is not allowed` anterior indica que la línea #5, columna #16 define el `any` tipo. Esto le ayuda a localizar el error.

Para evitar este problema, defina siempre el tipo de variable. Si no está seguro del tipo de variable, puede usar un [tipo de unión](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html). Esto puede ser útil para variables que tienen valores, que pueden ser de tipo , o (el tipo de valores `Range` es una unión de los `string` `number` `boolean` `Range` siguientes: `string | number | boolean` ).

### <a name="implicit-any"></a>Implícito `any`

Los tipos de variables typeScript se [pueden definir implícitamente.](https://www.typescriptlang.org/docs/handbook/type-inference.html) Si el compilador typeScript no puede determinar el tipo de una variable (ya sea porque el tipo no se define explícitamente o la inferencia de tipo no es posible), se trata de un error implícito y recibirá un error en tiempo de `any` compilación.

El caso más común en cualquier `any` implícito está en una declaración de variable, como `let value;` . Hay dos maneras de evitar esto:

* Asigne la variable a un tipo de identificación implícita ( `let value = 5;` o `let value = workbook.getWorksheet();` ).
* Escriba explícitamente la variable ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>No heredar clases o interfaces de Script de Office

Las clases e interfaces que se crean en el script de Office no pueden [extender ni implementar clases](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) o interfaces de scripts de Office. En otras palabras, nada en el espacio `ExcelScript` de nombres puede tener subclases o subinterfaces.

## <a name="incompatible-typescript-functions"></a>Funciones de TypeScript incompatibles

Las API de Scripts de Office no se pueden usar en lo siguiente:

* [Funciones de generador](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` no se admite

La función [de eval de](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript no se admite por motivos de seguridad.

## <a name="restricted-identifers"></a>Identifers restringidas

Las siguientes palabras no se pueden usar como identificadores en un script. Son términos reservados.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Solo funciones de flecha en devoluciones de llamada de matriz

Los scripts solo pueden usar [funciones de flecha](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) al proporcionar argumentos de devolución de llamada para los [métodos Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) No puede pasar ningún tipo de identificador o función "tradicional" a estos métodos.

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

El [linter](https://wikipedia.org/wiki/Lint_(software)) del Editor de código proporciona advertencias si el script puede tener problemas de rendimiento. Los casos y cómo trabajar alrededor de ellos se documentan en [Mejorar el rendimiento de los scripts de Office](web-client-performance.md).

## <a name="external-api-calls"></a>Llamadas de API externas

Para [obtener más información, vea Soporte](external-calls.md) técnico de llamadas de la API externa en Scripts de Office.

## <a name="see-also"></a>Vea también

* [Conceptos básicos de los Scripts de Office en Excel en la web](scripting-fundamentals.md)
* [Mejorar el rendimiento de los scripts de Office](web-client-performance.md)
