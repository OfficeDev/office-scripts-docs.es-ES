---
title: Restricciones de TypeScript en Office scripts
description: Los detalles del compilador TypeScript y linter usados por el editor de Office scripts.
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b5ba0dfe60081a0bb65dec4e694c7d534cb8df63
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585684"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Restricciones de TypeScript en Office scripts

Office scripts usan el lenguaje TypeScript. En su mayoría, cualquier código TypeScript o JavaScript funcionará en Office scripts. Sin embargo, el Editor de código aplica algunas restricciones para garantizar que el script funciona de forma coherente y según lo previsto con el Excel libro.

## <a name="no-any-type-in-office-scripts"></a>Ningún tipo de 'any' en Office scripts

Los [tipos](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) de escritura son opcionales en TypeScript, ya que los tipos se pueden deducir. Sin embargo, Office scripts requiere que una variable no pueda ser de [tipo ninguna](https://www.typescriptlang.org/docs/handbook/basic-types.html#any). Tanto explícitos como implícitos `any` no están permitidos en Office scripts. Estos casos se notifican como errores.

### <a name="explicit-any"></a>Explícito `any`

No se puede declarar explícitamente una variable como de tipo `any` en Office scripts (es decir, `let value: any;`). El `any` tipo provoca problemas al procesarlo Excel. Por ejemplo, es necesario `Range` saber que un valor es un `string`, `number`o `boolean`. Recibirá un error en tiempo de compilación (un error antes de ejecutar el script) si alguna variable se define explícitamente como `any` el tipo en el script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="El mensaje explícito &quot;any&quot; en el texto activado del Editor de código.":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Error explícito &quot;any&quot; en la ventana de la consola.":::

En la captura de pantalla anterior, `[2, 14] Explicit Any is not allowed` indica que la línea #2, columna #14 define el `any` tipo. Esto le ayuda a localizar el error.

Para evitar este problema, defina siempre el tipo de variable. Si no está seguro del tipo de variable, puede usar un tipo [de unión](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html). Esto puede ser útil para variables que `Range` tienen valores, `string`que pueden ser de tipo , `number`o `boolean` ( `Range` el tipo de valores es una unión de los siguientes: `string | number | boolean`).

### <a name="implicit-any"></a>Implícito `any`

Los tipos de variables typeScript se [pueden definir implícitamente](https://www.typescriptlang.org/docs/handbook/type-inference.html) . Si el compilador typeScript no puede determinar el tipo de una variable (ya sea porque el tipo no se define explícitamente o la inferencia de tipo no es posible), `any` se trata de un error implícito y recibirá un error en tiempo de compilación.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="El mensaje implícito &quot;any&quot; en el texto activado del Editor de código.":::

El caso más común en cualquier implícito `any` está en una declaración de variable, como `let value;`. Hay dos maneras de evitar esto:

* Asigne la variable a un tipo de identificación implícita (`let value = 5;` o `let value = workbook.getWorksheet();`).
* Escriba explícitamente la variable (`let value: number;`)

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>No se heredan Office o interfaces de script

Las clases e interfaces que se crean en su Office Script no pueden [extender ni implementar Office](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) clases o interfaces de scripts. En otras palabras, nada en el espacio `ExcelScript` de nombres puede tener subclases o subinterfaces.

## <a name="incompatible-typescript-functions"></a>Funciones de TypeScript incompatibles

Office API de scripts no se pueden usar en lo siguiente:

* [Funciones de generador](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` no se admite

La función [de eval de](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript no se admite por motivos de seguridad.

## <a name="restricted-identifiers"></a>Identificadores restringidos

Las siguientes palabras no se pueden usar como identificadores en un script. Son términos reservados.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Solo funciones de flecha en devoluciones de llamada de matriz

Los scripts solo pueden usar [funciones de flecha](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) al proporcionar argumentos de devolución de llamada para [los métodos Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) . No puede pasar ningún tipo de identificador o función "tradicional" a estos métodos.

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

## <a name="unions-of-excelscript-types-and-user-defined-types-arent-supported"></a>No se admiten `ExcelScript` uniones de tipos y tipos definidos por el usuario

Office scripts se convierten en tiempo de ejecución de bloques de código sincrónicos a asincrónicos. La comunicación con el libro a través [de las promesas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) está oculta del creador de scripts. Esta conversión no admite tipos de [unión](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) que incluyan `ExcelScript` tipos y tipos definidos por el usuario. En ese caso, se `Promise` devuelve al script, pero el compilador Office Script no lo espera y el creador de scripts no puede interactuar con el `Promise`.

En el ejemplo de código siguiente se muestra una unión no compatible entre `ExcelScript.Table` y una interfaz `MyTable` personalizada.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const selectedSheet = workbook.getActiveWorksheet();

  // This union is not supported.
  const tableOrMyTable: ExcelScript.Table | MyTable = selectedSheet.getTables()[0];

  // `getName` returns a promise that can't be resolved by the script.
  const name = tableOrMyTable.getName();

  // This logs "{}" instead of the table name.
  console.log(name);
}

interface MyTable {
  getName(): string
}
```

## <a name="constructors-dont-support-office-scripts-apis-and-console-statements"></a>Los constructores no admiten Office API y instrucciones `console` de scripts

`console`instrucciones y muchas API Office scripts requieren sincronización con el Excel de datos. Estas sincronizaciones usan instrucciones `await` en la versión en tiempo de ejecución compilada del script. `await` no se admite en [constructores](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Classes/constructor). Si necesita clases con constructores, evite usar Office API de scripts o `console` instrucciones en esos bloques de código.

En el ejemplo de código siguiente se muestra este escenario. Genera un error que indica `failed to load [code] [library]`.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  class MyClass {
    constructor() {
      // Console statements and Office Scripts APIs aren't supported in constructors.
      console.log("This won't print.");
    }
  }

  let test = new MyClass();
}
```

## <a name="performance-warnings"></a>Advertencias de rendimiento

El [linter](https://wikipedia.org/wiki/Lint_(software)) del Editor de código proporciona advertencias si el script puede tener problemas de rendimiento. Los casos y cómo trabajar alrededor de ellos se documentan en [Mejorar el rendimiento de los scripts de Office.](web-client-performance.md)

## <a name="external-api-calls"></a>Llamadas de API externas

Para [obtener más información, consulte Soporte de llamadas de la API externa en Office scripts](external-calls.md).

## <a name="see-also"></a>Vea también

* [Conceptos básicos de los Scripts de Office en Excel en la web](scripting-fundamentals.md)
* [Mejorar el rendimiento de los scripts de Office scripts](web-client-performance.md)
