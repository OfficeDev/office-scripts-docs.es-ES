---
title: Solución de problemas de scripts de Office
description: Sugerencias y técnicas de depuración de scripts de Office, así como recursos de ayuda.
ms.date: 10/30/2020
localization_priority: Normal
ms.openlocfilehash: b45957bd336edce527397253cacec8cb09df715a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 11/18/2020
ms.locfileid: "49342881"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="accb4-103">Solución de problemas de scripts de Office</span><span class="sxs-lookup"><span data-stu-id="accb4-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="accb4-104">Al desarrollar scripts de Office, puede cometer errores.</span><span class="sxs-lookup"><span data-stu-id="accb4-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="accb4-105">Es correcto.</span><span class="sxs-lookup"><span data-stu-id="accb4-105">It's okay.</span></span> <span data-ttu-id="accb4-106">Tenemos herramientas que ayudan a encontrar los problemas y que los scripts funcionan perfectamente.</span><span class="sxs-lookup"><span data-stu-id="accb4-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="accb4-107">Registros de la consola</span><span class="sxs-lookup"><span data-stu-id="accb4-107">Console logs</span></span>

<span data-ttu-id="accb4-108">En ocasiones, durante la solución de problemas, querrá imprimir los mensajes en la pantalla.</span><span class="sxs-lookup"><span data-stu-id="accb4-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="accb4-109">Estos pueden mostrar el valor actual de las variables o las rutas de código que se están desencadenando.</span><span class="sxs-lookup"><span data-stu-id="accb4-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="accb4-110">Para ello, registre el texto en la consola.</span><span class="sxs-lookup"><span data-stu-id="accb4-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="accb4-111">Las cadenas pasadas a `console.log` se mostrarán en la consola de registro del editor de código.</span><span class="sxs-lookup"><span data-stu-id="accb4-111">Strings passed to `console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="accb4-112">Para activar la consola, presione el botón de **puntos suspensivos** y seleccione **registros...**</span><span class="sxs-lookup"><span data-stu-id="accb4-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="accb4-113">Los registros no afectan al libro.</span><span class="sxs-lookup"><span data-stu-id="accb4-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="accb4-114">Mensajes de error</span><span class="sxs-lookup"><span data-stu-id="accb4-114">Error messages</span></span>

<span data-ttu-id="accb4-115">Cuando el script de Excel encuentra un problema en ejecución, produce un error.</span><span class="sxs-lookup"><span data-stu-id="accb4-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="accb4-116">Verá un mensaje emergente en el que se le preguntará si desea **ver los registros**.</span><span class="sxs-lookup"><span data-stu-id="accb4-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="accb4-117">Presione ese botón para abrir la consola y mostrar los errores.</span><span class="sxs-lookup"><span data-stu-id="accb4-117">Press that button to open the console and display any errors.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="accb4-118">La ficha automatizada no aparece o las secuencias de comandos de Office no están disponibles</span><span class="sxs-lookup"><span data-stu-id="accb4-118">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="accb4-119">Los pasos siguientes le ayudarán a solucionar los problemas relacionados con la ficha **automatizar** que no aparecen en Excel en la Web.</span><span class="sxs-lookup"><span data-stu-id="accb4-119">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="accb4-120">Asegúrese [de que su licencia de 365 de Microsoft incluye scripts de Office](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="accb4-120">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="accb4-121">[Compruebe que el explorador es compatible](platform-limits.md#browser-support).</span><span class="sxs-lookup"><span data-stu-id="accb4-121">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="accb4-122">[Asegúrese de que las cookies de terceros están habilitadas](platform-limits.md#third-party-cookies).</span><span class="sxs-lookup"><span data-stu-id="accb4-122">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="accb4-123">[Asegúrese de que su administrador no ha deshabilitado los scripts de Office en el centro de administración de Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="accb4-123">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a><span data-ttu-id="accb4-124">Recursos de ayuda</span><span class="sxs-lookup"><span data-stu-id="accb4-124">Help resources</span></span>

<span data-ttu-id="accb4-125">[Desbordamiento de pila](https://stackoverflow.com/questions/tagged/office-scripts) es una comunidad de desarrolladores que desea ayudar con los problemas de codificación.</span><span class="sxs-lookup"><span data-stu-id="accb4-125">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="accb4-126">A menudo, podrá encontrar la solución a su problema mediante una búsqueda rápida de desbordamiento de pila.</span><span class="sxs-lookup"><span data-stu-id="accb4-126">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="accb4-127">Si no es así, formule su pregunta y etiquete con la etiqueta "Office-scripts".</span><span class="sxs-lookup"><span data-stu-id="accb4-127">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="accb4-128">No olvide mencionar que está creando un *script* de Office, no un *complemento de* Office.</span><span class="sxs-lookup"><span data-stu-id="accb4-128">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="accb4-129">Si encuentra un problema con la API de JavaScript de Office, cree un problema en el repositorio de github [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) .</span><span class="sxs-lookup"><span data-stu-id="accb4-129">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="accb4-130">Los miembros del equipo de producto responderán a los problemas y proporcionarán asistencia.</span><span class="sxs-lookup"><span data-stu-id="accb4-130">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="accb4-131">La creación de un problema en el repositorio de **OfficeDev/Office-js** indica que ha encontrado un error en la biblioteca de la API de JavaScript de Office que el equipo del producto debe tratar.</span><span class="sxs-lookup"><span data-stu-id="accb4-131">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="accb4-132">Si hay un problema con el grabador de acciones o con el editor, envíe sus comentarios a través del botón **ayuda > comentarios** de Excel.</span><span class="sxs-lookup"><span data-stu-id="accb4-132">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="accb4-133">Consulte también</span><span class="sxs-lookup"><span data-stu-id="accb4-133">See also</span></span>

- [<span data-ttu-id="accb4-134">Scripts de Office en Excel en la Web</span><span class="sxs-lookup"><span data-stu-id="accb4-134">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="accb4-135">Conceptos básicos sobre el scripting de los scripts de Office en Excel en la web</span><span class="sxs-lookup"><span data-stu-id="accb4-135">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="accb4-136">Límites de plataforma con scripts de Office</span><span class="sxs-lookup"><span data-stu-id="accb4-136">Platform Limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="accb4-137">Mejorar el rendimiento de los scripts de Office</span><span class="sxs-lookup"><span data-stu-id="accb4-137">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="accb4-138">Deshacer los efectos de un script de Office</span><span class="sxs-lookup"><span data-stu-id="accb4-138">Undo the effects of an Office Script</span></span>](undo.md)
