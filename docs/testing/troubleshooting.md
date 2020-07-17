---
title: Solución de problemas de scripts de Office
description: Sugerencias y técnicas de depuración de scripts de Office, así como recursos de ayuda.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 6448980eec45214a589444229db0fd781b9fea13
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878622"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="04972-103">Solución de problemas de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="04972-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="04972-104">Al desarrollar scripts de Office, puede cometer errores.</span><span class="sxs-lookup"><span data-stu-id="04972-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="04972-105">Es correcto.</span><span class="sxs-lookup"><span data-stu-id="04972-105">It's okay.</span></span> <span data-ttu-id="04972-106">Tenemos herramientas que ayudan a encontrar los problemas y que los scripts funcionan perfectamente.</span><span class="sxs-lookup"><span data-stu-id="04972-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="04972-107">Registros de la consola</span><span class="sxs-lookup"><span data-stu-id="04972-107">Console logs</span></span>

<span data-ttu-id="04972-108">En ocasiones, durante la solución de problemas, querrá imprimir los mensajes en la pantalla.</span><span class="sxs-lookup"><span data-stu-id="04972-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="04972-109">Estos pueden mostrar el valor actual de las variables o las rutas de código que se están desencadenando.</span><span class="sxs-lookup"><span data-stu-id="04972-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="04972-110">Para ello, registre el texto en la consola.</span><span class="sxs-lookup"><span data-stu-id="04972-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="04972-111">Las cadenas pasadas a `console.log` se mostrarán en la consola de registro del editor de código.</span><span class="sxs-lookup"><span data-stu-id="04972-111">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="04972-112">Para activar la consola, presione el botón de **puntos suspensivos** y seleccione **registros...**</span><span class="sxs-lookup"><span data-stu-id="04972-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="04972-113">Los registros no afectan al libro.</span><span class="sxs-lookup"><span data-stu-id="04972-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="04972-114">Mensajes de error</span><span class="sxs-lookup"><span data-stu-id="04972-114">Error messages</span></span>

<span data-ttu-id="04972-115">Cuando el script de Excel encuentra un problema en ejecución, produce un error.</span><span class="sxs-lookup"><span data-stu-id="04972-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="04972-116">Verá un mensaje emergente en el que se le preguntará si desea **ver los registros**.</span><span class="sxs-lookup"><span data-stu-id="04972-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="04972-117">Presione ese botón para abrir la consola y mostrar los errores.</span><span class="sxs-lookup"><span data-stu-id="04972-117">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="04972-118">Recursos de ayuda</span><span class="sxs-lookup"><span data-stu-id="04972-118">Help resources</span></span>

<span data-ttu-id="04972-119">[Desbordamiento de pila](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores que desea ayudar con los problemas de codificación.</span><span class="sxs-lookup"><span data-stu-id="04972-119">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="04972-120">A menudo, podrá encontrar la solución a su problema mediante una búsqueda rápida de desbordamiento de pila.</span><span class="sxs-lookup"><span data-stu-id="04972-120">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="04972-121">Si no es así, formule su pregunta y etiquete con la etiqueta "Office-scripts".</span><span class="sxs-lookup"><span data-stu-id="04972-121">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="04972-122">No olvide mencionar que está creando un *script*de Office, no un *complemento de*Office.</span><span class="sxs-lookup"><span data-stu-id="04972-122">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="04972-123">Si encuentra un problema con la API de JavaScript de Office, cree un problema en el repositorio de github [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) .</span><span class="sxs-lookup"><span data-stu-id="04972-123">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="04972-124">Los miembros del equipo de producto responderán a los problemas y proporcionarán asistencia.</span><span class="sxs-lookup"><span data-stu-id="04972-124">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="04972-125">La creación de un problema en el repositorio de **OfficeDev/Office-js** indica que ha encontrado un error en la biblioteca de la API de JavaScript de Office que el equipo del producto debe tratar.</span><span class="sxs-lookup"><span data-stu-id="04972-125">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="04972-126">Si hay un problema con el grabador de acciones o con el editor, envíe sus comentarios a través del botón **ayuda > comentarios** de Excel.</span><span class="sxs-lookup"><span data-stu-id="04972-126">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="04972-127">Vea también</span><span class="sxs-lookup"><span data-stu-id="04972-127">See also</span></span>

- [<span data-ttu-id="04972-128">Scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="04972-128">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="04972-129">Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="04972-129">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="04972-130">Deshacer los efectos de un script de Office</span><span class="sxs-lookup"><span data-stu-id="04972-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="04972-131">Mejorar el rendimiento de los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="04972-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
