---
title: Restricciones de TypeScript en scripts de Office
description: Los detalles del compilador de TypeScript y el linter que usa el Editor de código de scripts de Office.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 41584ff23b333d17b2e267fdb3b0ec8741f3d203
ms.sourcegitcommit: df2b64603f91acb37bf95230efd538db0fbf9206
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 02/04/2021
ms.locfileid: "50099911"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="78361-103">Restricciones de TypeScript en scripts de Office</span><span class="sxs-lookup"><span data-stu-id="78361-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="78361-104">Los scripts de Office usan el lenguaje TypeScript.</span><span class="sxs-lookup"><span data-stu-id="78361-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="78361-105">En la mayoría de los casos, cualquier código TypeScript o JavaScript funcionará en un script de Office.</span><span class="sxs-lookup"><span data-stu-id="78361-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="78361-106">Sin embargo, hay algunas restricciones aplicadas por el Editor de código para garantizar que el script funciona de forma coherente y según lo previsto con el libro de Excel.</span><span class="sxs-lookup"><span data-stu-id="78361-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="78361-107">Ningún tipo "cualquiera" en scripts de Office</span><span class="sxs-lookup"><span data-stu-id="78361-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="78361-108">Escribir [tipos](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) es opcional en TypeScript, porque los tipos se pueden deducir.</span><span class="sxs-lookup"><span data-stu-id="78361-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="78361-109">Sin embargo, el script de Office requiere que una variable no pueda ser de [tipo alguno.](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="78361-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="78361-110">No se permiten explícitas `any` e implícitas en un script de Office.</span><span class="sxs-lookup"><span data-stu-id="78361-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="78361-111">Estos casos se notifican como errores.</span><span class="sxs-lookup"><span data-stu-id="78361-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="78361-112">Explicit `any`</span><span class="sxs-lookup"><span data-stu-id="78361-112">Explicit `any`</span></span>

<span data-ttu-id="78361-113">No puede declarar explícitamente una variable para que sea de tipo en scripts de `any` Office (es decir, `let someVariable: any;` ).</span><span class="sxs-lookup"><span data-stu-id="78361-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="78361-114">El `any` tipo provoca problemas cuando Excel lo procesa.</span><span class="sxs-lookup"><span data-stu-id="78361-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="78361-115">Por ejemplo, `Range` un necesita saber que un valor es un , o `string` `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="78361-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="78361-116">Recibirá un error en tiempo de compilación (un error antes de ejecutar el script) si alguna variable se define explícitamente como el tipo `any` en el script.</span><span class="sxs-lookup"><span data-stu-id="78361-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![El mensaje explícito de cualquier mensaje en el texto activando del editor de código](../images/explicit-any-editor-message.png)

![El error explícito en la ventana de la consola](../images/explicit-any-error-message.png)

<span data-ttu-id="78361-119">En la captura de pantalla `[5, 16] Explicit Any is not allowed` anterior, indica que la línea #5, la columna #16 el `any` tipo.</span><span class="sxs-lookup"><span data-stu-id="78361-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="78361-120">Esto le ayuda a localizar el error.</span><span class="sxs-lookup"><span data-stu-id="78361-120">This helps you locate the error.</span></span>

<span data-ttu-id="78361-121">Para evitar este problema, defina siempre el tipo de la variable.</span><span class="sxs-lookup"><span data-stu-id="78361-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="78361-122">Si no está seguro del tipo de variable, puede usar un tipo [de unión.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="78361-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="78361-123">Esto puede ser útil para variables que tienen valores, que pueden ser de tipo , o (el tipo de valores es una `Range` `string` unión de los `number` `boolean` `Range` siguientes: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="78361-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="78361-124">Implícito `any`</span><span class="sxs-lookup"><span data-stu-id="78361-124">Implicit `any`</span></span>

<span data-ttu-id="78361-125">Los tipos de variables de TypeScript pueden ser [implícitamente](( https://www.typescriptlang.org/docs/handbook/type-inference.html) definidos.</span><span class="sxs-lookup"><span data-stu-id="78361-125">TypeScript variable types can be [implicitly]((https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="78361-126">Si el compilador de TypeScript no puede determinar el tipo de una variable (ya sea porque el tipo no está definido explícitamente o la inferencia de tipo no es posible), es un error implícito y recibirá un error en tiempo de `any` compilación.</span><span class="sxs-lookup"><span data-stu-id="78361-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="78361-127">El caso más común en cualquier `any` implícito es en una declaración de variable, como `let value;` .</span><span class="sxs-lookup"><span data-stu-id="78361-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="78361-128">Hay dos formas de evitarlo:</span><span class="sxs-lookup"><span data-stu-id="78361-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="78361-129">Asignar la variable a un tipo de identificación implícita ( `let value = 5;` o `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="78361-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="78361-130">Escriba explícitamente la variable ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="78361-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="78361-131">No se heredan clases o interfaces de script de Office</span><span class="sxs-lookup"><span data-stu-id="78361-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="78361-132">Las clases e interfaces que se crean en el script de Office no pueden [ampliar ni](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) implementar clases o interfaces de scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="78361-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="78361-133">En otras palabras, nada en el `ExcelScript` espacio de nombres puede tener subclases o subinterfaces.</span><span class="sxs-lookup"><span data-stu-id="78361-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="78361-134">Funciones de TypeScript incompatibles</span><span class="sxs-lookup"><span data-stu-id="78361-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="78361-135">Las API de scripts de Office no se pueden usar en lo siguiente:</span><span class="sxs-lookup"><span data-stu-id="78361-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="78361-136">Funciones de generador</span><span class="sxs-lookup"><span data-stu-id="78361-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="78361-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="78361-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="78361-138">`eval` no se admite</span><span class="sxs-lookup"><span data-stu-id="78361-138">`eval` is not supported</span></span>

<span data-ttu-id="78361-139">La función de [eval de](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript no se admite por motivos de seguridad.</span><span class="sxs-lookup"><span data-stu-id="78361-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="78361-140">Identifers restringidos</span><span class="sxs-lookup"><span data-stu-id="78361-140">Restricted identifers</span></span>

<span data-ttu-id="78361-141">Las siguientes palabras no se pueden usar como identificadores en un script.</span><span class="sxs-lookup"><span data-stu-id="78361-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="78361-142">Son términos reservados.</span><span class="sxs-lookup"><span data-stu-id="78361-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a><span data-ttu-id="78361-143">Advertencias de rendimiento</span><span class="sxs-lookup"><span data-stu-id="78361-143">Performance warnings</span></span>

<span data-ttu-id="78361-144">El [linter](https://wikipedia.org/wiki/Lint_(software)) del Editor de código proporciona advertencias si el script puede tener problemas de rendimiento.</span><span class="sxs-lookup"><span data-stu-id="78361-144">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="78361-145">Los casos y cómo evitarlos se documentan en [Mejorar el rendimiento de los scripts de Office.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="78361-145">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="78361-146">Llamadas a API externas</span><span class="sxs-lookup"><span data-stu-id="78361-146">External API calls</span></span>

<span data-ttu-id="78361-147">Para [obtener más información, vea](external-calls.md) la compatibilidad con llamadas de API externas en scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="78361-147">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="78361-148">Vea también</span><span class="sxs-lookup"><span data-stu-id="78361-148">See also</span></span>

* [<span data-ttu-id="78361-149">Conceptos básicos de los Scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="78361-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="78361-150">Mejorar el rendimiento de los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="78361-150">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
