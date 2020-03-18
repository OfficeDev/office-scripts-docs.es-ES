---
title: Solución de problemas de scripts de Office
description: Sugerencias y técnicas de depuración de scripts de Office, así como recursos de ayuda.
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700356"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="e6c97-103">Solución de problemas de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="e6c97-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="e6c97-104">Al desarrollar scripts de Office, puede cometer errores.</span><span class="sxs-lookup"><span data-stu-id="e6c97-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="e6c97-105">Es correcto.</span><span class="sxs-lookup"><span data-stu-id="e6c97-105">It's okay.</span></span> <span data-ttu-id="e6c97-106">Tenemos herramientas que ayudan a encontrar los problemas y que los scripts funcionan perfectamente.</span><span class="sxs-lookup"><span data-stu-id="e6c97-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="e6c97-107">Registros de la consola</span><span class="sxs-lookup"><span data-stu-id="e6c97-107">Console logs</span></span>

<span data-ttu-id="e6c97-108">En ocasiones, durante la solución de problemas, querrá imprimir los mensajes en la pantalla.</span><span class="sxs-lookup"><span data-stu-id="e6c97-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="e6c97-109">Estos pueden mostrar el valor actual de las variables o las rutas de código que se están desencadenando.</span><span class="sxs-lookup"><span data-stu-id="e6c97-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="e6c97-110">Para ello, registre el texto en la consola.</span><span class="sxs-lookup"><span data-stu-id="e6c97-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> <span data-ttu-id="e6c97-111">No olvide los `load` datos de la `sync` hoja de cálculo y con el libro antes de registrar las propiedades del objeto.</span><span class="sxs-lookup"><span data-stu-id="e6c97-111">Don't forget to `load` worksheet data and `sync` with the workbook before logging object properties.</span></span>

<span data-ttu-id="e6c97-112">Las cadenas pasadas`console.log` a se mostrarán en la consola de registro del editor de código.</span><span class="sxs-lookup"><span data-stu-id="e6c97-112">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="e6c97-113">Para activar la consola, presione el botón de **puntos suspensivos** y seleccione **registros...**</span><span class="sxs-lookup"><span data-stu-id="e6c97-113">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="e6c97-114">Los registros no afectan al libro.</span><span class="sxs-lookup"><span data-stu-id="e6c97-114">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="e6c97-115">Mensajes de error</span><span class="sxs-lookup"><span data-stu-id="e6c97-115">Error messages</span></span>

<span data-ttu-id="e6c97-116">Cuando el script de Excel encuentra un problema en ejecución, produce un error.</span><span class="sxs-lookup"><span data-stu-id="e6c97-116">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="e6c97-117">Verá un mensaje emergente en el que se le preguntará si desea **ver los registros**.</span><span class="sxs-lookup"><span data-stu-id="e6c97-117">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="e6c97-118">Presione ese botón para abrir la consola y mostrar los errores.</span><span class="sxs-lookup"><span data-stu-id="e6c97-118">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="e6c97-119">Recursos de ayuda</span><span class="sxs-lookup"><span data-stu-id="e6c97-119">Help resources</span></span>

<span data-ttu-id="e6c97-120">[Desbordamiento de pila](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores que desea ayudar con los problemas de codificación.</span><span class="sxs-lookup"><span data-stu-id="e6c97-120">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="e6c97-121">A menudo, podrá encontrar la solución a su problema mediante una búsqueda rápida de desbordamiento de pila.</span><span class="sxs-lookup"><span data-stu-id="e6c97-121">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="e6c97-122">Si no es así, formule su pregunta y etiquete con la etiqueta "Office-scripts".</span><span class="sxs-lookup"><span data-stu-id="e6c97-122">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="e6c97-123">No olvide mencionar que está creando un *script*de Office, no un *complemento de*Office.</span><span class="sxs-lookup"><span data-stu-id="e6c97-123">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="e6c97-124">Si encuentra un problema con la API de JavaScript de Office, cree un problema en el repositorio de github [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) .</span><span class="sxs-lookup"><span data-stu-id="e6c97-124">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="e6c97-125">Los miembros del equipo de producto responderán a los problemas y proporcionarán asistencia.</span><span class="sxs-lookup"><span data-stu-id="e6c97-125">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="e6c97-126">La creación de un problema en el repositorio de **OfficeDev/Office-js** indica que ha encontrado un error en la biblioteca de la API de JavaScript de Office que el equipo del producto debe tratar.</span><span class="sxs-lookup"><span data-stu-id="e6c97-126">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="e6c97-127">Si hay un problema con el grabador de acciones o con el editor, envíe sus comentarios a través del botón **ayuda > comentarios** de Excel.</span><span class="sxs-lookup"><span data-stu-id="e6c97-127">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="e6c97-128">Vea también</span><span class="sxs-lookup"><span data-stu-id="e6c97-128">See also</span></span>

- [<span data-ttu-id="e6c97-129">Scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="e6c97-129">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="e6c97-130">Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="e6c97-130">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="e6c97-131">Deshacer los efectos de un script de Office</span><span class="sxs-lookup"><span data-stu-id="e6c97-131">Undo the effects of an Office Script</span></span>](undo.md)
