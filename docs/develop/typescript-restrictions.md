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
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="dd6d6-103">Restricciones de TypeScript en scripts de Office</span><span class="sxs-lookup"><span data-stu-id="dd6d6-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="dd6d6-104">Los scripts de Office usan el lenguaje TypeScript.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="dd6d6-105">En su mayoría, cualquier código TypeScript o JavaScript funcionará en un script de Office.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="dd6d6-106">Sin embargo, el Editor de código aplica algunas restricciones para garantizar que el script funciona de forma coherente y según lo previsto con el libro de Excel.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="dd6d6-107">No hay tipo 'any' en scripts de Office</span><span class="sxs-lookup"><span data-stu-id="dd6d6-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="dd6d6-108">Los [tipos](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) de escritura son opcionales en TypeScript, ya que los tipos se pueden deducir.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="dd6d6-109">Sin embargo, el script de Office requiere que una variable no pueda ser de [tipo .](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="dd6d6-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="dd6d6-110">No se permiten explícitas `any` ni implícitas en un script de Office.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="dd6d6-111">Estos casos se notifican como errores.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="dd6d6-112">Explícito `any`</span><span class="sxs-lookup"><span data-stu-id="dd6d6-112">Explicit `any`</span></span>

<span data-ttu-id="dd6d6-113">No puede declarar explícitamente una variable como de tipo en scripts de `any` Office (es decir, `let someVariable: any;` ).</span><span class="sxs-lookup"><span data-stu-id="dd6d6-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="dd6d6-114">El `any` tipo provoca problemas al procesarlo Excel.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="dd6d6-115">Por ejemplo, es `Range` necesario saber que un valor es un , o `string` `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="dd6d6-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="dd6d6-116">Recibirá un error en tiempo de compilación (un error antes de ejecutar el script) si alguna variable se define explícitamente como el tipo `any` en el script.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![El mensaje explícito de cualquier mensaje en el texto activado del editor de código](../images/explicit-any-editor-message.png)

![El error explícito de la ventana de la consola](../images/explicit-any-error-message.png)

<span data-ttu-id="dd6d6-119">En la captura de pantalla `[5, 16] Explicit Any is not allowed` anterior indica que la línea #5, columna #16 define el `any` tipo.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="dd6d6-120">Esto le ayuda a localizar el error.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-120">This helps you locate the error.</span></span>

<span data-ttu-id="dd6d6-121">Para evitar este problema, defina siempre el tipo de variable.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="dd6d6-122">Si no está seguro del tipo de variable, puede usar un [tipo de unión](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span><span class="sxs-lookup"><span data-stu-id="dd6d6-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="dd6d6-123">Esto puede ser útil para variables que tienen valores, que pueden ser de tipo , o (el tipo de valores `Range` es una unión de los `string` `number` `boolean` `Range` siguientes: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="dd6d6-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="dd6d6-124">Implícito `any`</span><span class="sxs-lookup"><span data-stu-id="dd6d6-124">Implicit `any`</span></span>

<span data-ttu-id="dd6d6-125">Los tipos de variables typeScript se [pueden definir implícitamente.](https://www.typescriptlang.org/docs/handbook/type-inference.html)</span><span class="sxs-lookup"><span data-stu-id="dd6d6-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="dd6d6-126">Si el compilador typeScript no puede determinar el tipo de una variable (ya sea porque el tipo no se define explícitamente o la inferencia de tipo no es posible), se trata de un error implícito y recibirá un error en tiempo de `any` compilación.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="dd6d6-127">El caso más común en cualquier `any` implícito está en una declaración de variable, como `let value;` .</span><span class="sxs-lookup"><span data-stu-id="dd6d6-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="dd6d6-128">Hay dos maneras de evitar esto:</span><span class="sxs-lookup"><span data-stu-id="dd6d6-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="dd6d6-129">Asigne la variable a un tipo de identificación implícita ( `let value = 5;` o `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="dd6d6-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="dd6d6-130">Escriba explícitamente la variable ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="dd6d6-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="dd6d6-131">No heredar clases o interfaces de Script de Office</span><span class="sxs-lookup"><span data-stu-id="dd6d6-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="dd6d6-132">Las clases e interfaces que se crean en el script de Office no pueden [extender ni implementar clases](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) o interfaces de scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="dd6d6-133">En otras palabras, nada en el espacio `ExcelScript` de nombres puede tener subclases o subinterfaces.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="dd6d6-134">Funciones de TypeScript incompatibles</span><span class="sxs-lookup"><span data-stu-id="dd6d6-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="dd6d6-135">Las API de Scripts de Office no se pueden usar en lo siguiente:</span><span class="sxs-lookup"><span data-stu-id="dd6d6-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="dd6d6-136">Funciones de generador</span><span class="sxs-lookup"><span data-stu-id="dd6d6-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="dd6d6-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="dd6d6-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="dd6d6-138">`eval` no se admite</span><span class="sxs-lookup"><span data-stu-id="dd6d6-138">`eval` is not supported</span></span>

<span data-ttu-id="dd6d6-139">La función [de eval de](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript no se admite por motivos de seguridad.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="dd6d6-140">Identifers restringidas</span><span class="sxs-lookup"><span data-stu-id="dd6d6-140">Restricted identifers</span></span>

<span data-ttu-id="dd6d6-141">Las siguientes palabras no se pueden usar como identificadores en un script.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="dd6d6-142">Son términos reservados.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="dd6d6-143">Solo funciones de flecha en devoluciones de llamada de matriz</span><span class="sxs-lookup"><span data-stu-id="dd6d6-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="dd6d6-144">Los scripts solo pueden usar [funciones de flecha](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) al proporcionar argumentos de devolución de llamada para los [métodos Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)</span><span class="sxs-lookup"><span data-stu-id="dd6d6-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="dd6d6-145">No puede pasar ningún tipo de identificador o función "tradicional" a estos métodos.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="dd6d6-146">Advertencias de rendimiento</span><span class="sxs-lookup"><span data-stu-id="dd6d6-146">Performance warnings</span></span>

<span data-ttu-id="dd6d6-147">El [linter](https://wikipedia.org/wiki/Lint_(software)) del Editor de código proporciona advertencias si el script puede tener problemas de rendimiento.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="dd6d6-148">Los casos y cómo trabajar alrededor de ellos se documentan en [Mejorar el rendimiento de los scripts de Office](web-client-performance.md).</span><span class="sxs-lookup"><span data-stu-id="dd6d6-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="dd6d6-149">Llamadas de API externas</span><span class="sxs-lookup"><span data-stu-id="dd6d6-149">External API calls</span></span>

<span data-ttu-id="dd6d6-150">Para [obtener más información, vea Soporte](external-calls.md) técnico de llamadas de la API externa en Scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="dd6d6-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="dd6d6-151">Vea también</span><span class="sxs-lookup"><span data-stu-id="dd6d6-151">See also</span></span>

* [<span data-ttu-id="dd6d6-152">Conceptos básicos de los Scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="dd6d6-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="dd6d6-153">Mejorar el rendimiento de los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="dd6d6-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
