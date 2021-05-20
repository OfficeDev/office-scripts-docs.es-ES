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
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="104a0-103">Restricciones de TypeScript en scripts de Office</span><span class="sxs-lookup"><span data-stu-id="104a0-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="104a0-104">Office Los scripts utilizan el lenguaje TypeScript.</span><span class="sxs-lookup"><span data-stu-id="104a0-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="104a0-105">En su mayor parte, cualquier código TypeScript o JavaScript funcionará en scripts Office.</span><span class="sxs-lookup"><span data-stu-id="104a0-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="104a0-106">Sin embargo, el Editor de código aplica algunas restricciones para garantizar que el script funcione de forma coherente y según lo previsto con el libro de trabajo Excel.</span><span class="sxs-lookup"><span data-stu-id="104a0-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="104a0-107">No hay ningún tipo de script de Office</span><span class="sxs-lookup"><span data-stu-id="104a0-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="104a0-108">Escribir [tipos](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) es opcional en TypeScript, porque los tipos se pueden inferir.</span><span class="sxs-lookup"><span data-stu-id="104a0-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="104a0-109">Sin embargo, Office Scripts requiere que una variable no pueda ser de [tipo .](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="104a0-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="104a0-110">Tanto explícito como implícito `any` no se permiten en Office scripts.</span><span class="sxs-lookup"><span data-stu-id="104a0-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="104a0-111">Estos casos se notifican como errores.</span><span class="sxs-lookup"><span data-stu-id="104a0-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="104a0-112">explícito `any`</span><span class="sxs-lookup"><span data-stu-id="104a0-112">Explicit `any`</span></span>

<span data-ttu-id="104a0-113">No puede declarar explícitamente que una variable sea de tipo `any` en scripts de Office (es decir, `let someVariable: any;` ).</span><span class="sxs-lookup"><span data-stu-id="104a0-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="104a0-114">El `any` tipo causa problemas cuando se procesa por Excel.</span><span class="sxs-lookup"><span data-stu-id="104a0-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="104a0-115">Por ejemplo, `Range` es necesario saber que un valor es un , o `string` `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="104a0-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="104a0-116">Recibirá un error en tiempo de compilación (un error antes de ejecutar el script) si alguna variable se define explícitamente como el `any` tipo del script.</span><span class="sxs-lookup"><span data-stu-id="104a0-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="El mensaje explícito de &quot;cualquiera&quot; en el texto flotante del Editor de código":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="El error explícito de &quot;cualquiera&quot; en la ventana de la consola":::

<span data-ttu-id="104a0-119">En la captura de pantalla anterior `[5, 16] Explicit Any is not allowed` indica que la línea #5, la columna #16 define el `any` tipo.</span><span class="sxs-lookup"><span data-stu-id="104a0-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="104a0-120">Esto le ayuda a localizar el error.</span><span class="sxs-lookup"><span data-stu-id="104a0-120">This helps you locate the error.</span></span>

<span data-ttu-id="104a0-121">Para evitar este problema, defina siempre el tipo de la variable.</span><span class="sxs-lookup"><span data-stu-id="104a0-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="104a0-122">Si no está seguro sobre el tipo de variable, puede utilizar un [tipo de unión](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span><span class="sxs-lookup"><span data-stu-id="104a0-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="104a0-123">Esto puede ser útil para las variables que mantienen `Range` valores, que pueden ser de tipo `string` , `number` o `boolean` (el tipo de valores es una unión `Range` de esos: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="104a0-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="104a0-124">implícito `any`</span><span class="sxs-lookup"><span data-stu-id="104a0-124">Implicit `any`</span></span>

<span data-ttu-id="104a0-125">Los tipos de variable TypeScript se pueden definir [implícitamente.](https://www.typescriptlang.org/docs/handbook/type-inference.html)</span><span class="sxs-lookup"><span data-stu-id="104a0-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="104a0-126">Si el compilador TypeScript no puede determinar el tipo de una variable (ya sea porque el tipo no se define explícitamente o la inferencia de tipo no es posible), entonces es un error implícito `any` y recibirá un error en tiempo de compilación.</span><span class="sxs-lookup"><span data-stu-id="104a0-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="104a0-127">El caso más común en cualquier implícito `any` está en una declaración de variable, como `let value;` .</span><span class="sxs-lookup"><span data-stu-id="104a0-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="104a0-128">Hay dos maneras de evitar esto:</span><span class="sxs-lookup"><span data-stu-id="104a0-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="104a0-129">Asigne la variable a un tipo implícitamente identificable ( `let value = 5;` o `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="104a0-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="104a0-130">Escriba explícitamente la variable ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="104a0-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="104a0-131">Sin heredar clases o interfaces de script Office</span><span class="sxs-lookup"><span data-stu-id="104a0-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="104a0-132">Las clases e interfaces que se crean en el script de Office no pueden [extender ni implementar](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office clases o interfaces de scripts.</span><span class="sxs-lookup"><span data-stu-id="104a0-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="104a0-133">En otras palabras, nada en el `ExcelScript` espacio de nombres puede tener subclases o subinterfaces.</span><span class="sxs-lookup"><span data-stu-id="104a0-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="104a0-134">Funciones incompatibles typescript</span><span class="sxs-lookup"><span data-stu-id="104a0-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="104a0-135">Office Las API de scripts no se pueden utilizar en lo siguiente:</span><span class="sxs-lookup"><span data-stu-id="104a0-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="104a0-136">Funciones del generador</span><span class="sxs-lookup"><span data-stu-id="104a0-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="104a0-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="104a0-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="104a0-138">`eval` no se admite</span><span class="sxs-lookup"><span data-stu-id="104a0-138">`eval` is not supported</span></span>

<span data-ttu-id="104a0-139">La [función eval de](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript no es compatible por razones de seguridad.</span><span class="sxs-lookup"><span data-stu-id="104a0-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="104a0-140">Identificadores restringidos</span><span class="sxs-lookup"><span data-stu-id="104a0-140">Restricted identifers</span></span>

<span data-ttu-id="104a0-141">Las siguientes palabras no se pueden usar como identificadores en un script.</span><span class="sxs-lookup"><span data-stu-id="104a0-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="104a0-142">Son términos reservados.</span><span class="sxs-lookup"><span data-stu-id="104a0-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="104a0-143">Solo funciones de flecha en devoluciones de llamada de matriz</span><span class="sxs-lookup"><span data-stu-id="104a0-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="104a0-144">Los scripts solo pueden usar [funciones de flecha](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) al proporcionar argumentos de devolución de llamada para los métodos [Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)</span><span class="sxs-lookup"><span data-stu-id="104a0-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="104a0-145">No puede pasar ningún tipo de identificador o función "tradicional" a estos métodos.</span><span class="sxs-lookup"><span data-stu-id="104a0-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="104a0-146">Advertencias de rendimiento</span><span class="sxs-lookup"><span data-stu-id="104a0-146">Performance warnings</span></span>

<span data-ttu-id="104a0-147">Linter del [](https://wikipedia.org/wiki/Lint_(software)) Editor de código proporciona advertencias si el script podría tener problemas de rendimiento.</span><span class="sxs-lookup"><span data-stu-id="104a0-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="104a0-148">Los casos y cómo solucionarlos se documentan en [Mejorar el rendimiento de los scripts de Office.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="104a0-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="104a0-149">Llamadas a la API externa</span><span class="sxs-lookup"><span data-stu-id="104a0-149">External API calls</span></span>

<span data-ttu-id="104a0-150">Consulte [Compatibilidad con llamadas a la API externa en scripts de Office](external-calls.md) para obtener más información.</span><span class="sxs-lookup"><span data-stu-id="104a0-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="104a0-151">Vea también</span><span class="sxs-lookup"><span data-stu-id="104a0-151">See also</span></span>

* [<span data-ttu-id="104a0-152">Conceptos básicos de los Scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="104a0-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="104a0-153">Mejore el rendimiento de sus scripts de Office</span><span class="sxs-lookup"><span data-stu-id="104a0-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
