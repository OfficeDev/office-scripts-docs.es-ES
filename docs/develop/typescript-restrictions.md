---
title: Restricciones de TypeScript en scripts de Office
description: Los detalles del compilador de TypeScript y el linter que usa el Editor de código de scripts de Office.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: d67e208561ce6ddd706d4c80cf29d2f013a32032
ms.sourcegitcommit: 98c7bc26f51dc8427669c571135c503d73bcee4c
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 02/06/2021
ms.locfileid: "50125937"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Restricciones de TypeScript en scripts de Office

Los scripts de Office usan el lenguaje TypeScript. En la mayoría de los casos, cualquier código TypeScript o JavaScript funcionará en un script de Office. Sin embargo, hay algunas restricciones aplicadas por el Editor de código para garantizar que el script funciona de forma coherente y según lo previsto con el libro de Excel.

## <a name="no-any-type-in-office-scripts"></a>Ningún tipo "cualquiera" en scripts de Office

Escribir [tipos](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) es opcional en TypeScript, porque los tipos se pueden deducir. Sin embargo, el script de Office requiere que una variable no pueda ser de [tipo alguno.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any) No se permiten explícitas `any` e implícitas en un script de Office. Estos casos se notifican como errores.

### <a name="explicit-any"></a>Explicit `any`

No puede declarar explícitamente una variable para que sea de tipo en scripts de `any` Office (es decir, `let someVariable: any;` ). El `any` tipo provoca problemas cuando Excel lo procesa. Por ejemplo, `Range` un necesita saber que un valor es un , o `string` `number` `boolean` . Recibirá un error en tiempo de compilación (un error antes de ejecutar el script) si alguna variable se define explícitamente como el tipo `any` en el script.

![El mensaje explícito de cualquier mensaje en el texto activando del editor de código](../images/explicit-any-editor-message.png)

![El error explícito en la ventana de la consola](../images/explicit-any-error-message.png)

En la captura de pantalla `[5, 16] Explicit Any is not allowed` anterior, indica que la línea #5, la columna #16 el `any` tipo. Esto le ayuda a localizar el error.

Para evitar este problema, defina siempre el tipo de la variable. Si no está seguro del tipo de variable, puede usar un tipo [de unión.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Esto puede ser útil para variables que tienen valores, que pueden ser de tipo , o (el tipo de valores es una `Range` `string` unión de los `number` `boolean` `Range` siguientes: `string | number | boolean` ).

### <a name="implicit-any"></a>Implícito `any`

Los tipos de variables de TypeScript se [pueden definir implícitamente.](https://www.typescriptlang.org/docs/handbook/type-inference.html) Si el compilador de TypeScript no puede determinar el tipo de una variable (ya sea porque el tipo no está definido explícitamente o la inferencia de tipo no es posible), es un error implícito y recibirá un error en tiempo de `any` compilación.

El caso más común en cualquier `any` implícito es en una declaración de variable, como `let value;` . Hay dos formas de evitarlo:

* Asignar la variable a un tipo de identificación implícita ( `let value = 5;` o `let value = workbook.getWorksheet();` ).
* Escriba explícitamente la variable ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>No se heredan clases o interfaces de script de Office

Las clases e interfaces que se crean en el script de Office no pueden [ampliar ni](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) implementar clases o interfaces de scripts de Office. En otras palabras, nada en el `ExcelScript` espacio de nombres puede tener subclases o subinterfaces.

## <a name="incompatible-typescript-functions"></a>Funciones de TypeScript incompatibles

Las API de scripts de Office no se pueden usar en lo siguiente:

* [Funciones de generador](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` no se admite

La función de [eval de](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript no se admite por motivos de seguridad.

## <a name="restricted-identifers"></a>Identifers restringidos

Las siguientes palabras no se pueden usar como identificadores en un script. Son términos reservados.

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a>Advertencias de rendimiento

El [linter](https://wikipedia.org/wiki/Lint_(software)) del Editor de código proporciona advertencias si el script puede tener problemas de rendimiento. Los casos y cómo evitarlos se documentan en [Mejorar el rendimiento de los scripts de Office.](web-client-performance.md)

## <a name="external-api-calls"></a>Llamadas a API externas

Para [obtener más información, vea](external-calls.md) la compatibilidad con llamadas de API externas en scripts de Office.

## <a name="see-also"></a>Vea también

* [Conceptos básicos de los Scripts de Office en Excel en la web](scripting-fundamentals.md)
* [Mejorar el rendimiento de los scripts de Office](web-client-performance.md)
