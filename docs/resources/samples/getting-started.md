---
title: Introducción a scripts de Office
description: Conceptos básicos sobre scripts de Office, incluidos los patrones de acceso, entorno y script.
ms.date: 04/01/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: 42b71a21470ac38535e3e95f091ec6267806e54a
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755059"
---
# <a name="getting-started"></a><span data-ttu-id="26336-103">Introducción</span><span class="sxs-lookup"><span data-stu-id="26336-103">Getting started</span></span>

<span data-ttu-id="26336-104">En esta sección se proporcionan detalles sobre los conceptos básicos de los scripts de Office, incluidos el acceso, el entorno, los conceptos básicos del script y algunos patrones de script básicos.</span><span class="sxs-lookup"><span data-stu-id="26336-104">This section provides details about the basics of Office Scripts including access, environment, script fundamentals, and few basic script patterns.</span></span>

## <a name="environment-setup"></a><span data-ttu-id="26336-105">Configuración del entorno</span><span class="sxs-lookup"><span data-stu-id="26336-105">Environment setup</span></span>

<span data-ttu-id="26336-106">Obtenga información sobre los conceptos básicos del editor de acceso, entorno y script.</span><span class="sxs-lookup"><span data-stu-id="26336-106">Learn about the basics of access, environment, and script editor.</span></span>

<span data-ttu-id="26336-107">[![Conceptos básicos de la aplicación scripts de Office](../../images/getting-started-env.png)](https://youtu.be/vvCtxsjPxo8 "Conceptos básicos de la aplicación scripts de Office")</span><span class="sxs-lookup"><span data-stu-id="26336-107">[![Basics of Office Scripts application](../../images/getting-started-env.png)](https://youtu.be/vvCtxsjPxo8 "Basics of Office Scripts application")</span></span>

### <a name="access"></a><span data-ttu-id="26336-108">Access</span><span class="sxs-lookup"><span data-stu-id="26336-108">Access</span></span>

<span data-ttu-id="26336-109">Scripts de Office requiere la configuración de administración disponible para el administrador de Microsoft 365 en **Configuración**  >  **de la organización Configuración Scripts**  >  **de Office**.</span><span class="sxs-lookup"><span data-stu-id="26336-109">Office Scripts requires admin settings available for Microsoft 365 administrator under **Settings** > **Org settings** > **Office Scripts**.</span></span> <span data-ttu-id="26336-110">De forma predeterminada, está activada para todos los usuarios.</span><span class="sxs-lookup"><span data-stu-id="26336-110">By default, it's turned on for all users.</span></span> <span data-ttu-id="26336-111">Hay dos configuraciones sub, que el administrador puede activar y desactivar.</span><span class="sxs-lookup"><span data-stu-id="26336-111">There are two sub-settings, which the admin can turn on and off.</span></span>

* <span data-ttu-id="26336-112">Capacidad para compartir scripts dentro de la organización</span><span class="sxs-lookup"><span data-stu-id="26336-112">Ability to share scripts within the organization</span></span>
* <span data-ttu-id="26336-113">Capacidad de usar scripts en Power Automate</span><span class="sxs-lookup"><span data-stu-id="26336-113">Ability to use scripts in Power Automate</span></span>

<span data-ttu-id="26336-114">Puede saber si tiene acceso a scripts de Office abriendo un archivo en Excel en la web (explorador) y viendo si la pestaña **Automatizar** aparece en la cinta de Opciones de Excel o no.</span><span class="sxs-lookup"><span data-stu-id="26336-114">You can tell if you have access to Office Scripts by opening a file in Excel on the web (browser) and seeing if the **Automate** tab appears in the Excel ribbon or not.</span></span>
<span data-ttu-id="26336-115">Si aún no puede ver la pestaña **Automatizar,** compruebe [esta sección de solución de problemas](../../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span><span class="sxs-lookup"><span data-stu-id="26336-115">If you still can't see the **Automate** tab, check [this troubleshooting section](../../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span></span>

### <a name="availability"></a><span data-ttu-id="26336-116">Disponibilidad</span><span class="sxs-lookup"><span data-stu-id="26336-116">Availability</span></span>

<span data-ttu-id="26336-117">Los scripts de Office solo están disponibles en Excel en la web para licencias de Enterprise E3+ (no se admiten cuentas de consumidor y E1).</span><span class="sxs-lookup"><span data-stu-id="26336-117">Office Scripts is available only in the Excel on the web for Enterprise E3+ licenses (Consumer and E1 accounts are not supported).</span></span> <span data-ttu-id="26336-118">Los scripts de Office aún no se admiten en Excel en Windows y Mac.</span><span class="sxs-lookup"><span data-stu-id="26336-118">Office Scripts is not yet supported in Excel on Windows and Mac.</span></span>

### <a name="scripts-and-editor"></a><span data-ttu-id="26336-119">Scripts y editor</span><span class="sxs-lookup"><span data-stu-id="26336-119">Scripts and editor</span></span>

<span data-ttu-id="26336-120">El editor de código se basa directamente en Excel en la web (versión en línea).</span><span class="sxs-lookup"><span data-stu-id="26336-120">The code editor is built right into Excel on the web (online version).</span></span> <span data-ttu-id="26336-121">Si has usado editores como Visual Studio Code o Sublime, esta experiencia de edición será bastante similar.</span><span class="sxs-lookup"><span data-stu-id="26336-121">If you have used editors like Visual Studio Code or Sublime, this editing experience will be quite similar.</span></span>
<span data-ttu-id="26336-122">La mayoría de las teclas de acceso directo que Visual Studio editor de código también usan trabajo en la experiencia de edición de scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="26336-122">Most of the shortcut keys that Visual Studio Code editor uses work in the Office Scripts editing experience as well.</span></span> <span data-ttu-id="26336-123">Consulte los siguientes volantes de teclas de método abreviado.</span><span class="sxs-lookup"><span data-stu-id="26336-123">Check out the following shortcut keys handouts.</span></span>

* [<span data-ttu-id="26336-124">macOS</span><span class="sxs-lookup"><span data-stu-id="26336-124">macOS</span></span>](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf)
* [<span data-ttu-id="26336-125">Windows</span><span class="sxs-lookup"><span data-stu-id="26336-125">Windows</span></span>](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf)

#### <a name="key-things-to-note"></a><span data-ttu-id="26336-126">Aspectos clave a tener en cuenta</span><span class="sxs-lookup"><span data-stu-id="26336-126">Key things to note</span></span>

* <span data-ttu-id="26336-127">Los scripts de Office solo están disponibles para los archivos almacenados en OneDrive para la Empresa, los sitios de SharePoint y los sitios de grupo.</span><span class="sxs-lookup"><span data-stu-id="26336-127">Office Scripts is only available for files stored in OneDrive for Business, SharePoint sites, and Team sites.</span></span>
* <span data-ttu-id="26336-128">El editor no muestra la extensión del script.</span><span class="sxs-lookup"><span data-stu-id="26336-128">The editor doesn't show the script's extension.</span></span> <span data-ttu-id="26336-129">En realidad, se trata de archivos TypeScript, pero se almacenan con una extensión personalizada denominada `.osts` .</span><span class="sxs-lookup"><span data-stu-id="26336-129">In reality, these are TypeScript files but they are stored with a custom extension called `.osts`.</span></span>
* <span data-ttu-id="26336-130">Los scripts se almacenan en su propia carpeta de OneDrive para la Empresa `My Files/Documents/OfficeScripts` .</span><span class="sxs-lookup"><span data-stu-id="26336-130">The scripts are stored in your own OneDrive for Business folder `My Files/Documents/OfficeScripts`.</span></span> <span data-ttu-id="26336-131">No necesitará administrar esta carpeta.</span><span class="sxs-lookup"><span data-stu-id="26336-131">You won't need to manage this folder.</span></span> <span data-ttu-id="26336-132">Por su parte, puede omitir este aspecto a medida que el editor administra la experiencia de visualización y edición.</span><span class="sxs-lookup"><span data-stu-id="26336-132">For your part, you can ignore this aspect as the editor manages the viewing/editing experience.</span></span>
* <span data-ttu-id="26336-133">Los scripts no se almacenan como parte de los archivos de Excel.</span><span class="sxs-lookup"><span data-stu-id="26336-133">Scripts are not stored as part of Excel files.</span></span> <span data-ttu-id="26336-134">Se almacenan por separado.</span><span class="sxs-lookup"><span data-stu-id="26336-134">They are stored separately.</span></span>
* <span data-ttu-id="26336-135">Puede compartir el script con un archivo de Excel que, en efecto, significa que está vinculando el script con el archivo, no adjuntarlo.</span><span class="sxs-lookup"><span data-stu-id="26336-135">You can share the script with an Excel file which in effect means you are linking the script with the file, not attaching it.</span></span> <span data-ttu-id="26336-136">Quien tenga acceso al archivo de Excel también podrá **ver,** **ejecutar** o hacer **una copia** del script.</span><span class="sxs-lookup"><span data-stu-id="26336-136">Whoever has access to the Excel file will also be able to **view**, **run**, or **make a copy** of the script.</span></span> <span data-ttu-id="26336-137">Esta es una diferencia clave en comparación con las macros de VBA.</span><span class="sxs-lookup"><span data-stu-id="26336-137">This is a key difference compared to VBA macros.</span></span>
* <span data-ttu-id="26336-138">A menos que comparta los scripts, nadie más podrá acceder a él, ya que reside en su propia biblioteca.</span><span class="sxs-lookup"><span data-stu-id="26336-138">Unless you share your scripts, no one else can access it as it resides in your own library.</span></span>
* <span data-ttu-id="26336-139">Los scripts no se pueden vincular desde un disco local ni desde ubicaciones de nube personalizadas.</span><span class="sxs-lookup"><span data-stu-id="26336-139">Scripts can't be linked from a local disk or custom cloud locations.</span></span> <span data-ttu-id="26336-140">Los scripts de Office solo reconocen y ejecutan un script que se encuentra en una ubicación predefinida (la carpeta de OneDrive mencionada anteriormente) o scripts compartidos.</span><span class="sxs-lookup"><span data-stu-id="26336-140">Office Scripts only recognizes and runs a script that is on predefined location (your OneDrive folder mentioned above) or shared scripts.</span></span>
* <span data-ttu-id="26336-141">Durante la edición, los archivos se guardan temporalmente en el explorador, pero tendrás que guardar el script antes de cerrar la ventana de Excel para guardarlo en la ubicación de OneDrive.</span><span class="sxs-lookup"><span data-stu-id="26336-141">During editing, files are temporarily saved in the browser but you'll have to save the script before closing the Excel window to save it to the OneDrive location.</span></span> <span data-ttu-id="26336-142">No olvide guardar el archivo después de editarlo.</span><span class="sxs-lookup"><span data-stu-id="26336-142">Don't forget to save the file after edits.</span></span>

## <a name="gentle-introduction-to-scripting"></a><span data-ttu-id="26336-143">Introducción suave a scripting</span><span class="sxs-lookup"><span data-stu-id="26336-143">Gentle introduction to scripting</span></span>

<span data-ttu-id="26336-144">Los scripts de Office son scripts independientes escritos en el lenguaje TypeScript que contienen instrucciones para realizar cierta automatización en el libro de Excel seleccionado.</span><span class="sxs-lookup"><span data-stu-id="26336-144">Office Scripts are standalone scripts written in the TypeScript language that contain instructions to perform some automation against the selected Excel workbook.</span></span> <span data-ttu-id="26336-145">Todas las instrucciones de automatización están independientes dentro de un script y los scripts no pueden invocar ni llamar a otros scripts.</span><span class="sxs-lookup"><span data-stu-id="26336-145">All automation instructions are self-contained within a script and scripts can't invoke or call other scripts.</span></span> <span data-ttu-id="26336-146">Todos los scripts se almacenan en archivos independientes y se almacenan en la carpeta de OneDrive del usuario.</span><span class="sxs-lookup"><span data-stu-id="26336-146">All scripts are stored in standalone files and stored on the user's OneDrive folder.</span></span> <span data-ttu-id="26336-147">Puede grabar un nuevo script, editar un script grabado o escribir un script nuevo completo desde cero, todo dentro de una interfaz de editor integrada.</span><span class="sxs-lookup"><span data-stu-id="26336-147">You can record a new script, edit a recorded script, or write a whole new script from scratch, all within a built-in editor interface.</span></span> <span data-ttu-id="26336-148">La mejor parte de los scripts de Office es que no necesitan ninguna configuración adicional de los usuarios.</span><span class="sxs-lookup"><span data-stu-id="26336-148">The best part of Office Scripts is that they don't need any further setup from users.</span></span> <span data-ttu-id="26336-149">No hay bibliotecas externas, páginas web o elementos de interfaz de usuario, configuración, etc. Todos los scripts de Office controlan toda la configuración del entorno y permiten un acceso fácil y rápido a la automatización a través de una interfaz de API sencilla.</span><span class="sxs-lookup"><span data-stu-id="26336-149">No external libraries, web pages, or UI elements, setup, etc. All the environment setup is handled by Office Scripts and it allows easy and fast access to automation through a simple API interface.</span></span>

<span data-ttu-id="26336-150">Algunos de los conceptos básicos útiles para comprender cómo editar y navegar por scripts son:</span><span class="sxs-lookup"><span data-stu-id="26336-150">Some of the basic concepts helpful to understand how to edit and navigate around scripts include:</span></span>

* <span data-ttu-id="26336-151">Sintaxis básica del lenguaje TypeScript</span><span class="sxs-lookup"><span data-stu-id="26336-151">Basic TypeScript language syntax</span></span>
* <span data-ttu-id="26336-152">Descripción de `main` la función y los argumentos</span><span class="sxs-lookup"><span data-stu-id="26336-152">Understanding of `main` function and arguments</span></span>
* <span data-ttu-id="26336-153">Objetos y jerarquía, métodos, propiedades</span><span class="sxs-lookup"><span data-stu-id="26336-153">Objects and hierarchy, methods, properties</span></span>
* <span data-ttu-id="26336-154">Colección (matriz): navegación y operaciones</span><span class="sxs-lookup"><span data-stu-id="26336-154">Collection (array): navigation and operations</span></span>
* <span data-ttu-id="26336-155">Definiciones de tipo</span><span class="sxs-lookup"><span data-stu-id="26336-155">Type definitions</span></span>
* <span data-ttu-id="26336-156">Entorno: registrar/editar, ejecutar, examinar resultados, compartir</span><span class="sxs-lookup"><span data-stu-id="26336-156">Environment: record/edit, run, examine results, share</span></span>

<span data-ttu-id="26336-157">En este vídeo y en la sección se explican algunos de estos conceptos en detalle.</span><span class="sxs-lookup"><span data-stu-id="26336-157">This video and section explain some of these concepts in detail.</span></span>

<span data-ttu-id="26336-158">[![Conceptos básicos de scripts de Office](../../images/getting-started-v_script.png)](https://youtu.be/8Zsrc1uaiiU "Conceptos básicos de scripts")</span><span class="sxs-lookup"><span data-stu-id="26336-158">[![Basics of Office Scripts](../../images/getting-started-v_script.png)](https://youtu.be/8Zsrc1uaiiU "Basics of Scripts")</span></span>

### <a name="language-typescript"></a><span data-ttu-id="26336-159">Idioma: TypeScript</span><span class="sxs-lookup"><span data-stu-id="26336-159">Language: TypeScript</span></span>

<span data-ttu-id="26336-160">[Los scripts](../../index.md) de Office se escriben con el lenguaje [TypeScript](https://www.typescriptlang.org/), que es un lenguaje de código abierto que se basa en JavaScript (uno de los más usados del mundo) agregando definiciones de tipos estáticos.</span><span class="sxs-lookup"><span data-stu-id="26336-160">[Office Scripts](../../index.md) is written using the [TypeScript language](https://www.typescriptlang.org/), which is an open-source language that builds on JavaScript (one of the world's most used) by adding static type definitions.</span></span> <span data-ttu-id="26336-161">Como dice el sitio web, proporcione una forma de describir la forma de un objeto, proporcionando mejor documentación y permitiendo que TypeScript valide que el código funciona `Types` correctamente.</span><span class="sxs-lookup"><span data-stu-id="26336-161">As the website says, `Types` provide a way to describe the shape of an object, providing better documentation, and allowing TypeScript to validate that your code is working correctly.</span></span>

<span data-ttu-id="26336-162">La sintaxis del lenguaje en sí se escribe con [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) con escrituras adicionales definidas en el script mediante convenciones de TypeScript.</span><span class="sxs-lookup"><span data-stu-id="26336-162">The language syntax itself is written using [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript) with additional typings defined in the script using TypeScript conventions.</span></span> <span data-ttu-id="26336-163">En su mayoría, puede pensar en scripts de Office como escritos en JavaScript.</span><span class="sxs-lookup"><span data-stu-id="26336-163">For the most part, you can think of Office Scripts as written in JavaScript.</span></span> <span data-ttu-id="26336-164">Es esencial que comprenda los conceptos básicos del lenguaje JavaScript para comenzar el recorrido de scripts de Office; aunque no es necesario ser competente para comenzar el recorrido de automatización.</span><span class="sxs-lookup"><span data-stu-id="26336-164">It is essential that you understand the basics of JavaScript language to begin your Office Scripts journey; though you don't need to be proficient at it to begin your automation journey.</span></span> <span data-ttu-id="26336-165">Con la grabadora de acciones de Scripts de Office, puede comprender las instrucciones de script porque se incluyen comentarios de código y puede seguir y realizar pequeñas modificaciones.</span><span class="sxs-lookup"><span data-stu-id="26336-165">With the Office Scripts' action recorder, you can understand the script statements because code comments are included and you can follow along and make small edits.</span></span>

<span data-ttu-id="26336-166">Las API de Scripts de Office, que permiten que el script interactúe con Excel, están diseñadas para usuarios finales que pueden no tener mucho fondo de codificación.</span><span class="sxs-lookup"><span data-stu-id="26336-166">Office Scripts APIs, which allow the script to interact with Excel, are designed for end-users who may not have much coding background.</span></span> <span data-ttu-id="26336-167">Las API se pueden invocar sincrónicamente y no es necesario conocer temas avanzados como promesas o devoluciones de llamada.</span><span class="sxs-lookup"><span data-stu-id="26336-167">APIs can be invoked synchronously and you don't need to know advanced topics such as promises or callbacks.</span></span> <span data-ttu-id="26336-168">El diseño de la API de scripts de Office proporciona:</span><span class="sxs-lookup"><span data-stu-id="26336-168">Office Scripts API design provides:</span></span>

* <span data-ttu-id="26336-169">Modelo de objetos simple con métodos, getters/setters.</span><span class="sxs-lookup"><span data-stu-id="26336-169">Simple object model with methods, getters/setters.</span></span>
* <span data-ttu-id="26336-170">Colecciones de objetos de fácil acceso como matrices regulares.</span><span class="sxs-lookup"><span data-stu-id="26336-170">Easy-to-access object collections as regular arrays.</span></span>
* <span data-ttu-id="26336-171">Opciones sencillas de control de errores.</span><span class="sxs-lookup"><span data-stu-id="26336-171">Simple error handling options.</span></span>
* <span data-ttu-id="26336-172">Rendimiento optimizado para escenarios selectos que ayudan a los usuarios a centrarse en el escenario disponible.</span><span class="sxs-lookup"><span data-stu-id="26336-172">Optimized performance for select scenarios helping users to focus on the scenario at hand.</span></span>

### <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="26336-173">`main` function: el punto inicial del script</span><span class="sxs-lookup"><span data-stu-id="26336-173">`main` function: The script's starting point</span></span>

<span data-ttu-id="26336-174">La ejecución de scripts de Office comienza en la `main` función.</span><span class="sxs-lookup"><span data-stu-id="26336-174">Office Scripts' execution begins at the `main` function.</span></span> <span data-ttu-id="26336-175">Un script es un único archivo que contiene una o varias funciones junto con declaraciones de tipos, interfaces, variables, etc. Para seguir con el script, comience con la función, ya que Excel siempre invoca primero `main` la función al ejecutar cualquier `main` script.</span><span class="sxs-lookup"><span data-stu-id="26336-175">A script is a single file containing one or many functions along with declarations of types, interfaces, variables, etc. To follow along with the script, begin with the `main` function as Excel always first invokes the `main` function when you execute any script.</span></span> <span data-ttu-id="26336-176">La función siempre tendrá al menos un argumento (o parámetro) denominado , que es solo un nombre de variable que identifica el libro actual con el que `main` `workbook` se ejecuta el script.</span><span class="sxs-lookup"><span data-stu-id="26336-176">The `main` function will always have at least one argument (or parameter) named `workbook`, which is just a variable name identifying the current workbook against which the script is running.</span></span> <span data-ttu-id="26336-177">Puede definir argumentos adicionales para su uso con la ejecución de Power Automate (sin conexión).</span><span class="sxs-lookup"><span data-stu-id="26336-177">You can define additional arguments for usage with Power Automate (offline) execution.</span></span>

* `function main(workbook: ExcelScript.Workbook)`

<span data-ttu-id="26336-178">Un script se puede organizar en funciones más pequeñas para ayudar con la reusabilidad del código, la claridad, etc. Otras funciones pueden estar dentro o fuera de la función principal, pero siempre en el mismo archivo.</span><span class="sxs-lookup"><span data-stu-id="26336-178">A script can be organized into smaller functions to aid with code reusability, clarity, etc. Other functions can be inside or outside of the main function but always in the same file.</span></span> <span data-ttu-id="26336-179">Un script es independiente y solo puede usar funciones definidas en el mismo archivo.</span><span class="sxs-lookup"><span data-stu-id="26336-179">A script is self-contained and can only use functions defined in the same file.</span></span> <span data-ttu-id="26336-180">Los scripts no pueden invocar ni llamar a otro script de Office.</span><span class="sxs-lookup"><span data-stu-id="26336-180">Scripts cannot invoke or call another Office Script.</span></span>

<span data-ttu-id="26336-181">Por lo tanto, en resumen:</span><span class="sxs-lookup"><span data-stu-id="26336-181">So, in summary:</span></span>

* <span data-ttu-id="26336-182">La `main` función es el punto de entrada de cualquier script.</span><span class="sxs-lookup"><span data-stu-id="26336-182">The `main` function is the entry point for any script.</span></span> <span data-ttu-id="26336-183">Cuando se ejecuta la función, la aplicación de Excel invoca esta función principal proporcionando el libro como su primer parámetro.</span><span class="sxs-lookup"><span data-stu-id="26336-183">When the function is executed, the Excel application invokes this main function by providing the workbook as its first parameter.</span></span>
* <span data-ttu-id="26336-184">Es importante mantener el primer argumento y `workbook` su declaración de tipo tal como aparece.</span><span class="sxs-lookup"><span data-stu-id="26336-184">It's important to keep the first argument `workbook` and its type declaration as it appears.</span></span> <span data-ttu-id="26336-185">Puede agregar nuevos argumentos a la función (consulte la siguiente sección) pero mantenga `main` el primer argumento tal y como está.</span><span class="sxs-lookup"><span data-stu-id="26336-185">You can add new arguments to the `main` function (see the next section) but do keep the first argument as is.</span></span>

:::image type="content" source="../../images/getting-started-main-introduction.png" alt-text="La función principal es el punto de entrada del script":::

#### <a name="send-or-receive-data-from-other-apps"></a><span data-ttu-id="26336-187">Enviar o recibir datos de otras aplicaciones</span><span class="sxs-lookup"><span data-stu-id="26336-187">Send or receive data from other apps</span></span>

<span data-ttu-id="26336-188">Puede conectar Excel a otras partes de su organización ejecutando scripts en [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="26336-188">You can connect Excel to other parts of your organization by running scripts in [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="26336-189">Obtenga más información sobre [cómo ejecutar scripts de Office en flujos de Power Automate.](../../develop/power-automate-integration.md)</span><span class="sxs-lookup"><span data-stu-id="26336-189">Learn more about [running Office Scripts in Power Automate flows](../../develop/power-automate-integration.md).</span></span>

<span data-ttu-id="26336-190">La forma de recibir o enviar datos desde y a Excel es a través de la `main` función.</span><span class="sxs-lookup"><span data-stu-id="26336-190">The way to receive or send data from and to Excel is through the `main` function.</span></span> <span data-ttu-id="26336-191">Piense en ella como la puerta de enlace de información que permite describir y usar los datos entrantes y salientes en el script.</span><span class="sxs-lookup"><span data-stu-id="26336-191">Think of it as the information gateway that allows incoming and outgoing data to be described and used in the script.</span></span> <span data-ttu-id="26336-192">Puede recibir datos de fuera del script mediante el tipo de datos y devolver cualquier dato reconocido por TypeScript, como , , o cualquier objeto en forma de interfaces que defina en el `string` `string` `number` `boolean` script.</span><span class="sxs-lookup"><span data-stu-id="26336-192">You can receive data from outside the script using the `string` data type and return any TypeScript-recognized data such as `string`, `number`, `boolean`, or any objects in the form of interfaces you define in the script.</span></span>

:::image type="content" source="../../images/getting-started-data-in-out.png" alt-text="Entradas y salidas de un script":::

#### <a name="use-functions-to-organize-and-reuse-code"></a><span data-ttu-id="26336-194">Usar funciones para organizar y reutilizar código</span><span class="sxs-lookup"><span data-stu-id="26336-194">Use functions to organize and reuse code</span></span>

<span data-ttu-id="26336-195">Puede usar funciones para organizar y reutilizar código dentro del script.</span><span class="sxs-lookup"><span data-stu-id="26336-195">You can use functions to organize and reuse code within your script.</span></span>

:::image type="content" source="../../images/getting-started-use-functions.png" alt-text="Uso de funciones en un script":::

### <a name="objects-hierarchy-methods-properties-collections"></a><span data-ttu-id="26336-197">Objetos, jerarquía, métodos, propiedades, colecciones</span><span class="sxs-lookup"><span data-stu-id="26336-197">Objects, hierarchy, methods, properties, collections</span></span>

<span data-ttu-id="26336-198">Todo el modelo de objetos de Excel se define en una estructura jerárquica de objetos, empezando por el objeto de libro de tipo `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="26336-198">All of Excel's object model is defined in a hierarchical structure of objects, beginning with the workbook object of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="26336-199">Un objeto puede contener métodos, propiedades y otros objetos dentro de él.</span><span class="sxs-lookup"><span data-stu-id="26336-199">An object can contain methods, properties, and other objects within it.</span></span> <span data-ttu-id="26336-200">Los objetos se vinculan entre sí mediante los métodos.</span><span class="sxs-lookup"><span data-stu-id="26336-200">Objects are linked to each other using the methods.</span></span> <span data-ttu-id="26336-201">El método de un objeto puede devolver otro objeto o colección de objetos.</span><span class="sxs-lookup"><span data-stu-id="26336-201">An object's method can return another object or collection of objects.</span></span> <span data-ttu-id="26336-202">El uso de la característica de IntelliSense (finalización de código) del editor de código es una excelente manera de explorar la jerarquía de objetos.</span><span class="sxs-lookup"><span data-stu-id="26336-202">Using the code editor's IntelliSense (code completion) feature is a great way to explore the object hierarchy.</span></span> <span data-ttu-id="26336-203">También puede usar el sitio de documentación [de referencia oficial](/javascript/api/office-scripts/overview) para seguir las relaciones entre objetos.</span><span class="sxs-lookup"><span data-stu-id="26336-203">You can also use the [official reference documentation site](/javascript/api/office-scripts/overview) to follow along with the relationships among objects.</span></span>

<span data-ttu-id="26336-204">Un [objeto](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Object) es una colección de propiedades y una propiedad es una asociación entre un nombre (o clave) y un valor.</span><span class="sxs-lookup"><span data-stu-id="26336-204">An [object](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Object) is a collection of properties, and a property is an association between a name (or key) and a value.</span></span> <span data-ttu-id="26336-205">El valor de una propiedad puede ser una función, en cuyo caso la propiedad se conoce como un método.</span><span class="sxs-lookup"><span data-stu-id="26336-205">A property's value can be a function, in which case the property is known as a method.</span></span> <span data-ttu-id="26336-206">En el caso del modelo de objetos scripts de Office, un objeto representa una cosa en el archivo de Excel con la que interactúan los usuarios, como un gráfico, un hipervínculo, una tabla dinámica, etc. También puede representar el comportamiento de un objeto, como los atributos de protección de una hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="26336-206">In the case of the Office Scripts object model, an object represents a thing in the Excel file that users interact with such as a chart, hyperlink, pivot-table, etc. It can also represent the behavior of an object such as the protection attributes of a worksheet.</span></span>

<span data-ttu-id="26336-207">El tema de los objetos y propiedades de TypeScript frente a los métodos es bastante profundo.</span><span class="sxs-lookup"><span data-stu-id="26336-207">The topic of TypeScript objects and properties vs methods is quite deep.</span></span> <span data-ttu-id="26336-208">Para empezar con el script y ser productivo, puede recordar algunas cosas básicas:</span><span class="sxs-lookup"><span data-stu-id="26336-208">In order to get started with the script and be productive, you can remember a few basic things:</span></span>

* <span data-ttu-id="26336-209">Se tiene acceso a ambos objetos y propiedades mediante notación (punto), con el objeto en el lado izquierdo de la propiedad o método `.` `.` en el lado derecho.</span><span class="sxs-lookup"><span data-stu-id="26336-209">Both objects and properties are accessed using `.` (dot) notation, with the object on the left side of the `.` and the property or method on the right side.</span></span> <span data-ttu-id="26336-210">Ejemplos: `hyperlink.address` , `range.getAddress()` .</span><span class="sxs-lookup"><span data-stu-id="26336-210">Examples: `hyperlink.address`, `range.getAddress()`.</span></span>
* <span data-ttu-id="26336-211">Las propiedades son escalares en la naturaleza (cadenas, booleanos, números).</span><span class="sxs-lookup"><span data-stu-id="26336-211">Properties are scalar in nature (strings, booleans, numbers).</span></span> <span data-ttu-id="26336-212">Por ejemplo, el nombre de un libro, la posición de una hoja de cálculo, el valor de si la tabla tiene un pie de página o no.</span><span class="sxs-lookup"><span data-stu-id="26336-212">For example, name of a workbook, position of a worksheet, the value of whether the table has a footer or not.</span></span>
* <span data-ttu-id="26336-213">Los métodos se "invocan" o se "ejecutan" con los paréntesis de apertura y cierre.</span><span class="sxs-lookup"><span data-stu-id="26336-213">Methods are 'invoked' or 'executed' using the open-close parentheses.</span></span> <span data-ttu-id="26336-214">Ejemplo: `table.delete()`.</span><span class="sxs-lookup"><span data-stu-id="26336-214">Example: `table.delete()`.</span></span> <span data-ttu-id="26336-215">A veces, un argumento se pasa a una función incluyéndolos entre paréntesis de apertura y cierre: `range.setValue('Hello')` .</span><span class="sxs-lookup"><span data-stu-id="26336-215">Sometimes an argument is passed to a function by including them between open-close parentheses: `range.setValue('Hello')`.</span></span> <span data-ttu-id="26336-216">Puede pasar muchos argumentos a una función (como se define en su contrato/firma) y separarlos mediante `,` .</span><span class="sxs-lookup"><span data-stu-id="26336-216">You can pass many arguments to a function (as defined by its contract/signature) and separate them using `,`.</span></span>  <span data-ttu-id="26336-217">Por ejemplo: `worksheet.addTable('A1:D6', true)`.</span><span class="sxs-lookup"><span data-stu-id="26336-217">For example: `worksheet.addTable('A1:D6', true)`.</span></span> <span data-ttu-id="26336-218">Puede pasar argumentos de cualquier tipo según lo requiera el método, como cadenas, números, booleanos o incluso otros objetos, por ejemplo, , donde es un objeto creado en otra parte del `worksheet.addTable(targetRange, true)` `targetRange` script.</span><span class="sxs-lookup"><span data-stu-id="26336-218">You can pass arguments of any type as required by the method such as strings, number, boolean, or even other objects, for example, `worksheet.addTable(targetRange, true)`, where `targetRange` is an object created elsewhere in the script.</span></span>
* <span data-ttu-id="26336-219">Los métodos pueden devolver algo como una propiedad escalar (nombre, dirección, etc.) u otro objeto (intervalo, gráfico) o no devolver nada en absoluto (como el caso con `delete` métodos).</span><span class="sxs-lookup"><span data-stu-id="26336-219">Methods can return a thing such as a scalar property (name, address, etc.) or another object (range, chart), or not return anything at all (such as the case with `delete` methods).</span></span> <span data-ttu-id="26336-220">Recibirá lo que devuelve el método declarando una variable o asignando a una variable existente.</span><span class="sxs-lookup"><span data-stu-id="26336-220">You receive what the method returns by declaring a variable or assigning to an existing variable.</span></span> <span data-ttu-id="26336-221">Puede ver que en el lado izquierdo de la instrucción, como `const table = worksheet.addTable('A1:D6', true)` .</span><span class="sxs-lookup"><span data-stu-id="26336-221">You can see that on the left hand side of statement such as `const table = worksheet.addTable('A1:D6', true)`.</span></span>
* <span data-ttu-id="26336-222">En su mayoría, el modelo de objetos scripts de Office consta de objetos con métodos que vinculan varias partes del modelo de objetos de Excel.</span><span class="sxs-lookup"><span data-stu-id="26336-222">For the most part, the Office Scripts object model consists of objects with methods that link various parts of the Excel object model.</span></span> <span data-ttu-id="26336-223">Muy rara vez se encontrará con propiedades que sean de valores escalares u objetos.</span><span class="sxs-lookup"><span data-stu-id="26336-223">Very rarely you'll come across properties that are of scalar or object values.</span></span>
* <span data-ttu-id="26336-224">En scripts de Office, un método de modelo de objetos de Excel debe contener paréntesis de apertura y cierre.</span><span class="sxs-lookup"><span data-stu-id="26336-224">In Office Scripts, an Excel object model method has to contain open-close parentheses.</span></span> <span data-ttu-id="26336-225">No se permite el uso de métodos sin ellos (como asignar un método a una variable).</span><span class="sxs-lookup"><span data-stu-id="26336-225">Using methods without them is not allowed (such as assigning a method to a variable).</span></span>

<span data-ttu-id="26336-226">Veamos algunos métodos en el `workbook` objeto.</span><span class="sxs-lookup"><span data-stu-id="26336-226">Let's look at a few methods on the `workbook` object.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Return a boolean (true or false) setting of whether the workbook is set to auto-save or not. 
    const autoSave = workbook.getAutoSave(); 
    // Get workbook name.
    const name = workbook.getName();
    // Get active cell range object.
    const cell = workbook.getActiveCell();
    // Get table named SALES.
    const cell = workbook.getTable('SALES');
    // Get all slicer objects.
    const slicers = workbook.getSlicers();
}
```

<span data-ttu-id="26336-227">En este ejemplo:</span><span class="sxs-lookup"><span data-stu-id="26336-227">In this example:</span></span>

* <span data-ttu-id="26336-228">Los métodos del `workbook` objeto, como `getAutoSave()` y `getName()` devuelven una propiedad escalar (cadena, número, booleano).</span><span class="sxs-lookup"><span data-stu-id="26336-228">The methods of the `workbook` object such as `getAutoSave()` and `getName()` return a scalar property (string, number, boolean).</span></span>
* <span data-ttu-id="26336-229">Métodos como `getActiveCell()` devolver otro objeto.</span><span class="sxs-lookup"><span data-stu-id="26336-229">Methods such as `getActiveCell()` return another object.</span></span>
* <span data-ttu-id="26336-230">El método acepta un argumento (nombre de tabla en este caso) y `getTable()` devuelve una tabla específica en el libro.</span><span class="sxs-lookup"><span data-stu-id="26336-230">The `getTable()` method accepts an argument (table name in this case) and returns a specific table in the workbook.</span></span>
* <span data-ttu-id="26336-231">El `getSlicers()` método devuelve una matriz (a la que se hace referencia en muchos lugares como una colección) de todos los objetos de segmentación de datos del libro.</span><span class="sxs-lookup"><span data-stu-id="26336-231">The `getSlicers()` method returns an array (referred to in many places as a collection) of all slicer objects within the workbook.</span></span>

<span data-ttu-id="26336-232">Observará que todos estos métodos tienen un prefijo, que es solo una convención usada en el modelo de objetos scripts de Office para transmitir que el método `get` devuelve algo.</span><span class="sxs-lookup"><span data-stu-id="26336-232">You'll notice that all of these methods have a `get` prefix, which is just a convention used in the Office Scripts object model to convey that the method is returning something.</span></span> <span data-ttu-id="26336-233">También se les conoce comúnmente como "getters".</span><span class="sxs-lookup"><span data-stu-id="26336-233">They are also commonly referred to as 'getters'.</span></span>

<span data-ttu-id="26336-234">Hay otros dos tipos de métodos que veremos en el siguiente ejemplo:</span><span class="sxs-lookup"><span data-stu-id="26336-234">There are two other types of methods that we'll now see in the next example:</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get a worksheet named 'Sheet1.
    const sheet = workbook.getWorksheet('Sheet1'); 
    // Set name to SALES.
    sheet.setName('SALES');
    // Position the worksheet at the beginning.
    sheet.setPosition(0);
}
```

<span data-ttu-id="26336-235">En este ejemplo:</span><span class="sxs-lookup"><span data-stu-id="26336-235">In this example:</span></span>

* <span data-ttu-id="26336-236">El `setName()` método establece un nuevo nombre en la hoja de cálculo.</span><span class="sxs-lookup"><span data-stu-id="26336-236">The `setName()` method sets a new name to the worksheet.</span></span> <span data-ttu-id="26336-237">`setPosition()` establece la posición en la primera celda.</span><span class="sxs-lookup"><span data-stu-id="26336-237">`setPosition()` sets the position to the first cell.</span></span>
* <span data-ttu-id="26336-238">Estos métodos modifican el archivo de Excel estableciendo una propiedad o comportamiento del libro.</span><span class="sxs-lookup"><span data-stu-id="26336-238">Such methods modify the Excel file by setting a property or behavior of the workbook.</span></span> <span data-ttu-id="26336-239">Estos métodos se denominan "setters".</span><span class="sxs-lookup"><span data-stu-id="26336-239">These methods are called 'setters'.</span></span>
* <span data-ttu-id="26336-240">Normalmente, los "setters" tienen un "getter" complementario, por ejemplo, y , ambos `worksheet.getPosition` `worksheet.setPosition` son métodos.</span><span class="sxs-lookup"><span data-stu-id="26336-240">Typically 'setters' have a companion 'getter', for example, `worksheet.getPosition` and `worksheet.setPosition`, both of which are methods.</span></span>

#### <a name="undefined-and-null-primitive-types"></a><span data-ttu-id="26336-241">`undefined` y `null` tipos primitivos</span><span class="sxs-lookup"><span data-stu-id="26336-241">`undefined` and `null` primitive types</span></span>

<span data-ttu-id="26336-242">Los siguientes son dos tipos de datos primitivos que debe tener en cuenta:</span><span class="sxs-lookup"><span data-stu-id="26336-242">The following are two primitive data types that you must be aware of:</span></span>

1. <span data-ttu-id="26336-243">El valor [`null`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/null) representa la ausencia intencionada de cualquier valor de objeto.</span><span class="sxs-lookup"><span data-stu-id="26336-243">The value [`null`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/null) represents the intentional absence of any object value.</span></span> <span data-ttu-id="26336-244">Es uno de los valores primitivos de JavaScript y se usa para indicar que una variable no tiene ningún valor.</span><span class="sxs-lookup"><span data-stu-id="26336-244">It is one of JavaScript's primitive values and is used to indicate that a variable has no value.</span></span>
1. <span data-ttu-id="26336-245">Una variable a la que no se ha asignado un valor es de tipo [`undefined`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/undefined) .</span><span class="sxs-lookup"><span data-stu-id="26336-245">A variable that has not been assigned a value is of type [`undefined`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/undefined).</span></span> <span data-ttu-id="26336-246">Un método o instrucción también puede devolver si la variable que se está `undefined` evaluando no tiene un valor asignado.</span><span class="sxs-lookup"><span data-stu-id="26336-246">A method or statement can also return `undefined` if the variable that's being evaluated doesn't have an assigned value.</span></span>

<span data-ttu-id="26336-247">Estos dos tipos se recortan como parte del tratamiento de errores y pueden causar bastante dolor de cabeza si no se manejan correctamente.</span><span class="sxs-lookup"><span data-stu-id="26336-247">These two types crop up as part of error handling and can cause quite a bit of headache if not handled properly.</span></span> <span data-ttu-id="26336-248">Afortunadamente, TypeScript/JavaScript ofrece una forma de comprobar si una variable es de tipo `undefined` o `null` .</span><span class="sxs-lookup"><span data-stu-id="26336-248">Fortunately, TypeScript/JavaScript offers a way to check if a variable is of type `undefined` or `null`.</span></span> <span data-ttu-id="26336-249">Hablaremos sobre algunas de esas comprobaciones en secciones posteriores, incluido el control de errores.</span><span class="sxs-lookup"><span data-stu-id="26336-249">We will talk about some of those checks in later sections, including error handling.</span></span>

#### <a name="method-chaining"></a><span data-ttu-id="26336-250">Encadenamiento de métodos</span><span class="sxs-lookup"><span data-stu-id="26336-250">Method chaining</span></span>

<span data-ttu-id="26336-251">Puede usar la notación de puntos para conectar los objetos que se devuelven desde un método para acortar el código.</span><span class="sxs-lookup"><span data-stu-id="26336-251">You can use dot notation to connect objects being returned from a method to shorten your code.</span></span> <span data-ttu-id="26336-252">A veces, esta técnica facilita la lectura y administración del código.</span><span class="sxs-lookup"><span data-stu-id="26336-252">Sometimes this technique makes the code easy to read and manage.</span></span> <span data-ttu-id="26336-253">Sin embargo, hay pocas cosas que tener en cuenta.</span><span class="sxs-lookup"><span data-stu-id="26336-253">However, there are few things to be aware of.</span></span> <span data-ttu-id="26336-254">Veamos los ejemplos siguientes.</span><span class="sxs-lookup"><span data-stu-id="26336-254">Let's look at the following examples.</span></span>

<span data-ttu-id="26336-255">El siguiente código obtiene la celda activa y la siguiente celda y, a continuación, establece el valor.</span><span class="sxs-lookup"><span data-stu-id="26336-255">The following code gets the active cell and the next cell, then sets the value.</span></span> <span data-ttu-id="26336-256">Este es un buen candidato para usar el encadenamiento, ya que este código se realizará correctamente todo el tiempo.</span><span class="sxs-lookup"><span data-stu-id="26336-256">This is a good candidate to use chaining as this code will succeed all the time.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getActiveCell().getOffsetRange(0,1).setValue('Next cell');
}
```

<span data-ttu-id="26336-257">Sin embargo, el siguiente código (que obtiene una tabla denominada **SALES** y activa su estilo de columna con bandas) tiene un problema.</span><span class="sxs-lookup"><span data-stu-id="26336-257">However, the following code (which gets a table named **SALES** and turns on its banded column style) has an issue.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  workbook.getTable('SALES').setShowBandedColumns(true);
}
```

<span data-ttu-id="26336-258">¿Qué ocurre **si la tabla SALES** no existe?</span><span class="sxs-lookup"><span data-stu-id="26336-258">What if the **SALES** table doesn't exist?</span></span> <span data-ttu-id="26336-259">El script producirá un error (se muestra a continuación) porque devuelve (que es un tipo de JavaScript que indica que no hay ninguna tabla `getTable('SALES')` `undefined` como **SALES**).</span><span class="sxs-lookup"><span data-stu-id="26336-259">The script will fail with an error (shown next) because `getTable('SALES')` returns `undefined` (which is a JavaScript type indicating that there is no table such as **SALES**).</span></span> <span data-ttu-id="26336-260">Llamar al `setShowBandedColumns` método on no tiene sentido, es decir, y, por lo tanto, el `undefined` script termina en un `undefined.setShowBandedColumns(true)` error.</span><span class="sxs-lookup"><span data-stu-id="26336-260">Calling the `setShowBandedColumns` method on `undefined` makes no sense, that is, `undefined.setShowBandedColumns(true)`, and hence the script ends in an error.</span></span>

```text
Line 2: Cannot read property 'setShowBandedColumns' of undefined
```

<span data-ttu-id="26336-261">Puede usar [](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Optional_chaining) el operador de encadenamiento opcional que proporciona una forma de simplificar el acceso a valores a través de objetos conectados cuando sea posible que una referencia o método sea o (que es la forma de JavaScript de indicar un objeto o resultado sin asignación o inexistente) para controlar esta `undefined` `null` condición.</span><span class="sxs-lookup"><span data-stu-id="26336-261">You could use the [optional chaining operator](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Optional_chaining) that provides a way to simplify accessing values through connected objects when it's possible that a reference or method may be `undefined` or `null` (which is JavaScript's way of indicating an unassigned or nonexistent object or result) to handle this condition.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // This line will not fail as the setShowBandedColumns method is executed only if the SALES table is present.
    workbook.getTable('SALES')?.setShowBandedColumns(true); 
}
```

<span data-ttu-id="26336-262">Si desea controlar las condiciones de objeto inexistentes o el tipo que devuelve un método, es mejor asignar el valor devuelto del método y controlarlo `undefined` por separado.</span><span class="sxs-lookup"><span data-stu-id="26336-262">If you wish to handle nonexistent object conditions or `undefined` type being returned by a method, then it is better to assign the return value from the method and handle that separately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const salesTable = workbook.getTable('SALES');
    if (salesTable) {
        salesTable.setShowBandedColumns(true);
    } else { 
        // Handle this condition.
    }
}
```

#### <a name="get-object-reference"></a><span data-ttu-id="26336-263">Obtener referencia de objeto</span><span class="sxs-lookup"><span data-stu-id="26336-263">Get object reference</span></span>

<span data-ttu-id="26336-264">El `workbook` objeto se le entrega en la `main` función.</span><span class="sxs-lookup"><span data-stu-id="26336-264">The `workbook` object is given to you in the `main` function.</span></span> <span data-ttu-id="26336-265">Puede empezar a usar el `workbook` objeto y acceder a sus métodos directamente.</span><span class="sxs-lookup"><span data-stu-id="26336-265">You can begin to use the `workbook` object and access its methods directly.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get workbook name.
    const name = workbook.getName();
    // Display name to console.
    console.log(name);
}
```

<span data-ttu-id="26336-266">Para usar todos los demás objetos dentro del libro, comience con el objeto y vaya hacia abajo en la jerarquía hasta que llegue al `workbook` objeto que está buscando.</span><span class="sxs-lookup"><span data-stu-id="26336-266">For using all other objects within the workbook, begin with `workbook` object and go down the hierarchy until you get to the object you are looking for.</span></span> <span data-ttu-id="26336-267">Puede obtener la referencia de objeto mediante la captura del objeto mediante su método o recuperando `get` la colección de objetos como se muestra a continuación:</span><span class="sxs-lookup"><span data-stu-id="26336-267">You can get the object reference by fetching the object using its `get` method or by retrieving the collection of objects as shown below:</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    const sheet = workbook.getActiveWorksheet();
    // Fetch using an ID or key.
    const sheet = workbook.getWorksheet('SomeSheetName');
    // Invoke methods on the object.
    sheet.setPosition(0); 
    
    // Get collection of methods.
    const tables = sheet.getTables();
    console.log('Total tables in this sheet: ' + tables.length);
}
```

#### <a name="check-if-an-object-exists-then-delete-and-add"></a><span data-ttu-id="26336-268">Comprobar si existe un objeto, eliminar y agregar</span><span class="sxs-lookup"><span data-stu-id="26336-268">Check if an object exists, then delete, and add</span></span>

<span data-ttu-id="26336-269">Para crear un objeto, por ejemplo, con un nombre predefinido, siempre es mejor quitar un objeto similar que pueda existir y, a continuación, agregarlo.</span><span class="sxs-lookup"><span data-stu-id="26336-269">For creating an object, say with a predefined name, it is always better to remove a similar object that may exist and then add it.</span></span> <span data-ttu-id="26336-270">Puede hacerlo con el siguiente patrón.</span><span class="sxs-lookup"><span data-stu-id="26336-270">You can do that using the following pattern.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added. 
  let name = "Index";
  // Check if the worksheet already exists. If not, add the worksheet.
  let sheet = workbook.getWorksheet('Index');
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Call the delete method on the object to remove it. 
    sheet.delete();
  } 
    // Add a blank worksheet. 
  console.log(`Adding the worksheet named  ${name}.`)
  const indexSheet = workbook.addWorksheet("Index");
}

```

<span data-ttu-id="26336-271">Como alternativa, para eliminar un objeto que puede existir o no, use el siguiente patrón.</span><span class="sxs-lookup"><span data-stu-id="26336-271">Alternatively, for deleting an object that may or may not exist, use the following pattern.</span></span>

```TypeScript
    // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
    workbook.getWorksheet('Index')?.delete(); 
```

#### <a name="note-about-adding-an-object"></a><span data-ttu-id="26336-272">Nota sobre cómo agregar un objeto</span><span class="sxs-lookup"><span data-stu-id="26336-272">Note about adding an object</span></span>

<span data-ttu-id="26336-273">Para crear, insertar o agregar un objeto como una segmentación de datos, una tabla dinámica, una hoja de cálculo, etc., use el método **add_Object_** correspondiente.</span><span class="sxs-lookup"><span data-stu-id="26336-273">To create, insert, or add an object such as a slicer, pivot table, worksheet, etc., use the corresponding **add_Object_** method.</span></span> <span data-ttu-id="26336-274">Este método está disponible en su objeto primario.</span><span class="sxs-lookup"><span data-stu-id="26336-274">Such a method is available on its parent object.</span></span> <span data-ttu-id="26336-275">Por ejemplo, el `addChart()` método está disponible en el `worksheet` objeto.</span><span class="sxs-lookup"><span data-stu-id="26336-275">For example, the `addChart()` method is available on `worksheet` object.</span></span> <span data-ttu-id="26336-276">El **add_Object_** devuelve el objeto que crea.</span><span class="sxs-lookup"><span data-stu-id="26336-276">The **add_Object_** method returns the object it creates.</span></span> <span data-ttu-id="26336-277">Reciba el valor devuelto y úselo más adelante en el script.</span><span class="sxs-lookup"><span data-stu-id="26336-277">Receive the returned value and use it later in your script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Add object and get a reference to it. 
  const indexSheet = workbook.addWorksheet("Index");
  // Use it elsewhere in the script 
  console.log(indexSheet.getPosition());
}

```

<span data-ttu-id="26336-278">Como alternativa, para eliminar un objeto que puede existir o no, use este patrón:</span><span class="sxs-lookup"><span data-stu-id="26336-278">Alternatively, for deleting an object that may or may not exist, use this pattern:</span></span>

```TypeScript
    workbook.getWorksheet('Index')?.delete(); // The ? preceding delete() will ensure that the API is only invoked if the object exists. 
```

#### <a name="collections"></a><span data-ttu-id="26336-279">Colecciones</span><span class="sxs-lookup"><span data-stu-id="26336-279">Collections</span></span>

<span data-ttu-id="26336-280">Las colecciones son objetos como tablas, gráficos, columnas, etc. que se pueden recuperar como una matriz e iterar para su procesamiento.</span><span class="sxs-lookup"><span data-stu-id="26336-280">Collections are objects such as tables, charts, columns, etc. that can be retrieved as an array and iterated over for processing.</span></span> <span data-ttu-id="26336-281">Puede recuperar una colección con el método correspondiente y procesar los datos en un bucle mediante una de las muchas técnicas de recorrido de matriz `get` de TypeScript, como:</span><span class="sxs-lookup"><span data-stu-id="26336-281">You can retrieve a collection using the corresponding `get` method and process the data in a loop using one of many TypeScript array traversal techniques such as:</span></span>

* [<span data-ttu-id="26336-282">`for` o `while`</span><span class="sxs-lookup"><span data-stu-id="26336-282">`for` or `while`</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
* [`for..of`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/for...of)
* [`forEach`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/forEach)

* [<span data-ttu-id="26336-283">Conceptos básicos de idioma de las matrices</span><span class="sxs-lookup"><span data-stu-id="26336-283">Language basics of arrays</span></span>](https://developer.mozilla.org//docs/Learn/JavaScript/First_steps/Arrays)

<span data-ttu-id="26336-284">Este script muestra cómo usar colecciones admitidas en las API de Scripts de Office.</span><span class="sxs-lookup"><span data-stu-id="26336-284">This script demonstrates how to use collections supported in Office Scripts APIs.</span></span> <span data-ttu-id="26336-285">Colore cada ficha de hoja de cálculo del archivo con un color aleatorio.</span><span class="sxs-lookup"><span data-stu-id="26336-285">It colors each worksheet tab in the file with a random color.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get all sheets as a collection.
  const sheets = workbook.getWorksheets();
  const names = sheets.map ((sheet) => sheet.getName());
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  // Get information from specific sheets within the collection.
  console.log(`First sheet name is: ${names[0]}`);
  if (sheets.length > 1) {
    console.log(`Last sheet's Id is: ${sheets[sheets.length -1].getId()}`);
  }
  // Color each worksheet with random color.
  for (const sheet of sheets) {
    sheet.setTabColor(`#${Math.random().toString(16).substr(-6)}`);
  }
}
```

## <a name="type-declarations"></a><span data-ttu-id="26336-286">Declaraciones de tipo</span><span class="sxs-lookup"><span data-stu-id="26336-286">Type declarations</span></span>

<span data-ttu-id="26336-287">Las declaraciones de tipo ayudan a los usuarios a comprender el tipo de variable con la que están tratando.</span><span class="sxs-lookup"><span data-stu-id="26336-287">Type declarations help users understand the type of variable they are dealing with.</span></span> <span data-ttu-id="26336-288">Ayuda con la finalización automática de métodos y ayuda en las comprobaciones de calidad del tiempo de desarrollo.</span><span class="sxs-lookup"><span data-stu-id="26336-288">It helps with auto-completion of methods and assists in development time quality checks.</span></span>

<span data-ttu-id="26336-289">Puede encontrar declaraciones de tipo en el script en varios lugares, como la declaración de función, la declaración de variables, IntelliSense definiciones, etc.</span><span class="sxs-lookup"><span data-stu-id="26336-289">You can find type declarations in the script in various places including function declaration, variable declaration, IntelliSense definitions, etc.</span></span>

<span data-ttu-id="26336-290">Ejemplos:</span><span class="sxs-lookup"><span data-stu-id="26336-290">Examples:</span></span>

* `function main(workbook: ExcelScript.Workbook)`
* `let myRange: ExcelScript.Range;`
* `function getMaxAmount(range: ExcelScript.Range): number`

<span data-ttu-id="26336-291">Puede identificar los tipos fácilmente en el editor de código, ya que normalmente aparece de forma distinta en un color diferente.</span><span class="sxs-lookup"><span data-stu-id="26336-291">You can identify the types easily in the code editor as it usually appears distinctly in a different color.</span></span> <span data-ttu-id="26336-292">Normalmente, `:` dos puntos precede a la declaración de tipo.</span><span class="sxs-lookup"><span data-stu-id="26336-292">A colon `:` usually precedes the type declaration.</span></span>  

<span data-ttu-id="26336-293">Los tipos de escritura pueden ser opcionales en TypeScript porque la inferencia de tipos le permite obtener mucha potencia sin escribir código adicional.</span><span class="sxs-lookup"><span data-stu-id="26336-293">Writing types can be optional in TypeScript because type inference allows you to get a lot of power without writing additional code.</span></span> <span data-ttu-id="26336-294">En su mayoría, el lenguaje TypeScript es bueno para deducir los tipos de variables.</span><span class="sxs-lookup"><span data-stu-id="26336-294">For the most part, the TypeScript language is good at inferring the types of variables.</span></span> <span data-ttu-id="26336-295">Sin embargo, en algunos casos, los scripts de Office requieren que las declaraciones de tipo se definan explícitamente si el idioma no puede identificar claramente el tipo.</span><span class="sxs-lookup"><span data-stu-id="26336-295">However, in certain cases, Office Scripts require the type declarations to be explicitly defined if the language is unable to clearly identify the type.</span></span> <span data-ttu-id="26336-296">Además, no se permite explícita `any` o implícita en el script de Office.</span><span class="sxs-lookup"><span data-stu-id="26336-296">Also, explicit or implicit `any` is not allowed in Office Script.</span></span> <span data-ttu-id="26336-297">Más información más adelante.</span><span class="sxs-lookup"><span data-stu-id="26336-297">More on that later.</span></span>

### <a name="excelscript-types"></a><span data-ttu-id="26336-298">`ExcelScript` tipos</span><span class="sxs-lookup"><span data-stu-id="26336-298">`ExcelScript` types</span></span>

<span data-ttu-id="26336-299">En Scripts de Office, usará los siguientes tipos de tipos.</span><span class="sxs-lookup"><span data-stu-id="26336-299">In Office Scripts, you will use the following kinds of types.</span></span>

* <span data-ttu-id="26336-300">Tipos de idioma nativo como `number` , , , , , `string` `object` `boolean` `null` etc.</span><span class="sxs-lookup"><span data-stu-id="26336-300">Native language types such as `number`, `string`, `object`, `boolean`, `null`, etc.</span></span>
* <span data-ttu-id="26336-301">Tipos de API de Excel.</span><span class="sxs-lookup"><span data-stu-id="26336-301">Excel API types.</span></span> <span data-ttu-id="26336-302">Empiezan por `ExcelScript` .</span><span class="sxs-lookup"><span data-stu-id="26336-302">They begin with `ExcelScript`.</span></span> <span data-ttu-id="26336-303">Por ejemplo, `ExcelScript.Range` , `ExcelScript.Table` , etc.</span><span class="sxs-lookup"><span data-stu-id="26336-303">For example, `ExcelScript.Range`, `ExcelScript.Table`, etc.</span></span>
* <span data-ttu-id="26336-304">Cualquier interfaz personalizada que haya definido en el script mediante `interface` instrucciones.</span><span class="sxs-lookup"><span data-stu-id="26336-304">Any custom interfaces you may have defined in the script using `interface` statements.</span></span>

<span data-ttu-id="26336-305">Vea los ejemplos de cada uno de estos grupos a continuación.</span><span class="sxs-lookup"><span data-stu-id="26336-305">See examples of each of these groups next.</span></span>

<span data-ttu-id="26336-306">**_Tipos de idioma nativo_**</span><span class="sxs-lookup"><span data-stu-id="26336-306">**_Native language types_**</span></span>

<span data-ttu-id="26336-307">En el ejemplo siguiente, observe los lugares `string` donde , y se han `number` `boolean` usado.</span><span class="sxs-lookup"><span data-stu-id="26336-307">In the following example, notice places where `string`, `number`, and `boolean` have been used.</span></span> <span data-ttu-id="26336-308">Estos son tipos de **lenguaje TypeScript** nativos.</span><span class="sxs-lookup"><span data-stu-id="26336-308">These are native **TypeScript** language types.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook)
{
  const table = workbook.getActiveWorksheet().getTables()[0];
  const sales = table.getColumnByName('Sales').getRange().getValues();
  console.log(sales);
  // Add 100 to each value.
  const revisedSales = salesAs1DArray.map(data => data as number + 100);
  // Add a column.
  table.addColumn(-1, revisedSales);  
}
/**
 * Extract a column from 2D array and return result.
 */
function extractColumn(data: (string | number | boolean)[][], index: number): (string | number | boolean)[] {

  const column = data.map((row) => {
    return row[index];
  })
  return column;
}
/**
 * Convert a flat array into a 2D array that can be used as range column.
 */
function convertColumnTo2D(data: (string | number | boolean)[]): (string | number | boolean)[][] {

  const columnAs2D = data.map((row) => {
    return [row];
  })
  return columnAs2D;
}
```

<span data-ttu-id="26336-309">**_Tipos de ExcelScript_**</span><span class="sxs-lookup"><span data-stu-id="26336-309">**_ExcelScript types_**</span></span>

<span data-ttu-id="26336-310">En el ejemplo siguiente, una función auxiliar toma dos argumentos.</span><span class="sxs-lookup"><span data-stu-id="26336-310">In the following example, a helper function takes two arguments.</span></span> <span data-ttu-id="26336-311">La primera es la `sheet` variable que es de tipo `ExcelScript.Worksheet` type.</span><span class="sxs-lookup"><span data-stu-id="26336-311">The first one is the `sheet` variable which is of type `ExcelScript.Worksheet` type.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for update.
    if (usedRange) { 
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);      
    targetRange.setValues([data]);
    return;
}
```

<span data-ttu-id="26336-312">**_Tipos personalizados_**</span><span class="sxs-lookup"><span data-stu-id="26336-312">**_Custom types_**</span></span>

<span data-ttu-id="26336-313">La interfaz personalizada `ReportImages` se usa para devolver imágenes a otra acción de flujo.</span><span class="sxs-lookup"><span data-stu-id="26336-313">The custom interface `ReportImages` is used to return images to another flow action.</span></span> <span data-ttu-id="26336-314">La declaración de función incluye instrucciones para decir a TypeScript que se devuelve un `main` `: ReportImages` objeto de ese tipo.</span><span class="sxs-lookup"><span data-stu-id="26336-314">The `main` function declaration includes `: ReportImages` instruction to tell TypeScript that an object of that type is being returned.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  let chart = workbook.getWorksheet("Sheet1").getCharts()[0];
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  
  const chartImage = chart.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

### <a name="type-assertion-overriding-the-type"></a><span data-ttu-id="26336-315">Aserción de tipo (invalidación del tipo)</span><span class="sxs-lookup"><span data-stu-id="26336-315">Type assertion (overriding the type)</span></span>

<span data-ttu-id="26336-316">Como indica la documentación [de](https://www.typescriptlang.org/docs/handbook/basic-types.html#type-assertions) TypeScript, "A veces, terminas en una situación en la que sabrás más sobre un valor que TypeScript.</span><span class="sxs-lookup"><span data-stu-id="26336-316">As the TypeScript [documentation](https://www.typescriptlang.org/docs/handbook/basic-types.html#type-assertions) states, "Sometimes you'll end up in a situation where you'll know more about a value than TypeScript does.</span></span> <span data-ttu-id="26336-317">Normalmente, esto ocurrirá cuando sepa que el tipo de alguna entidad podría ser más específico que su tipo actual.</span><span class="sxs-lookup"><span data-stu-id="26336-317">Usually, this will happen when you know the type of some entity could be more specific than its current type.</span></span> <span data-ttu-id="26336-318">Las aserciones de tipo son una forma de decir al compilador "confía en mí, sé lo que estoy haciendo".</span><span class="sxs-lookup"><span data-stu-id="26336-318">Type assertions are a way to tell the compiler “trust me, I know what I'm doing.”</span></span> <span data-ttu-id="26336-319">Una aserción de tipo es como un tipo que se convierte en otros idiomas, pero no realiza ninguna comprobación o reestructuración especial de los datos.</span><span class="sxs-lookup"><span data-stu-id="26336-319">A type assertion is like a type cast in other languages, but it performs no special checking or restructuring of data.</span></span> <span data-ttu-id="26336-320">No tiene ningún impacto en tiempo de ejecución y el compilador lo usa exclusivamente".</span><span class="sxs-lookup"><span data-stu-id="26336-320">It has no runtime impact and is used purely by the compiler."</span></span>

<span data-ttu-id="26336-321">Puede afirmar el tipo con la palabra `as` clave o con corchetes angulares, como se muestra en el código siguiente.</span><span class="sxs-lookup"><span data-stu-id="26336-321">You can assert the type using the `as` keyword or using angle brackets as shown in following code.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let data = workbook.getActiveCell().getValue();
  // Since the add10 function only accepts number, assert data's type as number, otherwise the script cannot be run.
  const answer1 = add10(data as number);
  const answer2 = add10(<number> data);
}

function add10(data: number) { 
  return data + 10;
}
```

#### <a name="any-type-in-the-script"></a><span data-ttu-id="26336-322">Tipo 'any' en el script</span><span class="sxs-lookup"><span data-stu-id="26336-322">'any' type in the script</span></span>

<span data-ttu-id="26336-323">El [sitio web de TypeScript indica:](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)</span><span class="sxs-lookup"><span data-stu-id="26336-323">The [TypeScript website states](https://www.typescriptlang.org/docs/handbook/basic-types.html#any):</span></span>

  <span data-ttu-id="26336-324">En algunas situaciones, no toda la información de tipo está disponible o su declaración requiere una cantidad de esfuerzo inapropiada.</span><span class="sxs-lookup"><span data-stu-id="26336-324">In some situations, not all type information is available or its declaration would take an inappropriate amount of effort.</span></span> <span data-ttu-id="26336-325">Estos pueden ocurrir para los valores del código que se ha escrito sin TypeScript o una biblioteca de terceros.</span><span class="sxs-lookup"><span data-stu-id="26336-325">These may occur for values from code that has been written without TypeScript or a 3rd party library.</span></span> <span data-ttu-id="26336-326">En estos casos, es posible que deseemos no participar en la comprobación de tipo.</span><span class="sxs-lookup"><span data-stu-id="26336-326">In these cases, we might want to opt-out of type checking.</span></span> <span data-ttu-id="26336-327">Para ello, etiquetamos estos valores con el `any` tipo:</span><span class="sxs-lookup"><span data-stu-id="26336-327">To do so, we label these values with the `any` type:</span></span>

  ```TypeScript
  declare function getValue(key: string): any;
  // OK, return value of 'getValue' is not checked
  const str: string = getValue("myString");
  ```

<span data-ttu-id="26336-328">**Explicit `any` is NOT allowed**</span><span class="sxs-lookup"><span data-stu-id="26336-328">**Explicit `any` is NOT allowed**</span></span>

```TypeScript
// This is not allowed
let someVariable: any; 
```

<span data-ttu-id="26336-329">El `any` tipo presenta desafíos a la forma en que Scripts de Office procesa las API de Excel.</span><span class="sxs-lookup"><span data-stu-id="26336-329">The `any` type presents challenges to the way Office Scripts processes the Excel APIs.</span></span> <span data-ttu-id="26336-330">Provoca problemas cuando las variables se envían a las API de Excel para su procesamiento.</span><span class="sxs-lookup"><span data-stu-id="26336-330">It causes issues when the variables are sent to Excel APIs for processing.</span></span> <span data-ttu-id="26336-331">Conocer el tipo de variables usadas en el script es esencial para el procesamiento del script y, por lo tanto, se prohíbe la definición explícita de cualquier variable `any` con tipo.</span><span class="sxs-lookup"><span data-stu-id="26336-331">Knowing the type of variables used in the script is essential to the processing of script and hence explicit definition of any variable with `any` type is prohibited.</span></span> <span data-ttu-id="26336-332">Recibirá un error en tiempo de compilación (error antes de ejecutar el script) si hay alguna variable con `any` el tipo declarado en el script.</span><span class="sxs-lookup"><span data-stu-id="26336-332">You will receive a compile-time error (error prior to running the script) if there is any variable with `any` type declared in the script.</span></span> <span data-ttu-id="26336-333">También verá un error en el editor.</span><span class="sxs-lookup"><span data-stu-id="26336-333">You will see an error in the editor as well.</span></span>

:::image type="content" source="../../images/getting-started-eanyi.png" alt-text="Error explícito &quot;cualquiera&quot;":::

:::image type="content" source="../../images/getting-started-expany.png" alt-text="Error explícito &quot;cualquiera&quot; que se muestra en Salida":::

<span data-ttu-id="26336-336">En el código mostrado en la imagen anterior, indica que la `[5, 16] Explicit Any is not allowed` línea 5 columna 16 declara el `any` tipo.</span><span class="sxs-lookup"><span data-stu-id="26336-336">In the code displayed in the previous image, `[5, 16] Explicit Any is not allowed` indicates that line 5 column 16 declares the `any` type.</span></span> <span data-ttu-id="26336-337">Esto le ayuda a buscar la línea de código que contiene el error.</span><span class="sxs-lookup"><span data-stu-id="26336-337">This helps you locate the line of code that contains the error.</span></span>

<span data-ttu-id="26336-338">Para evitar este problema, declare siempre el tipo de la variable.</span><span class="sxs-lookup"><span data-stu-id="26336-338">To get around this issue, always declare the type of the variable.</span></span>

<span data-ttu-id="26336-339">Si no está seguro del tipo de variable, un truco interesante en TypeScript le permite definir [tipos de unión](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span><span class="sxs-lookup"><span data-stu-id="26336-339">If you are uncertain about the type of a variable, one cool trick in TypeScript allows you to define [union types](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="26336-340">Esto se puede usar para que las variables puedan contener valores de intervalo, que pueden ser de muchos tipos.</span><span class="sxs-lookup"><span data-stu-id="26336-340">This can be used for variables to hold a range values, which can be of many types.</span></span>

```TypeScript
// Define value as a union type rather than 'any' type.
let value: (string | number | boolean);
value = someValue_from_another_source;
//...
someRange.setValue(value);
```

### <a name="type-inference"></a><span data-ttu-id="26336-341">Inferencia de tipo</span><span class="sxs-lookup"><span data-stu-id="26336-341">Type inference</span></span>

<span data-ttu-id="26336-342">En TypeScript, hay varios lugares en los que se usa la [inferencia](https://www.typescriptlang.org/docs/handbook/type-inference.html) de tipos para proporcionar información de tipo cuando no hay ninguna anotación de tipo explícita.</span><span class="sxs-lookup"><span data-stu-id="26336-342">In TypeScript, there are several places where [type inference](https://www.typescriptlang.org/docs/handbook/type-inference.html) is used to provide type information when there is no explicit type annotation.</span></span> <span data-ttu-id="26336-343">Por ejemplo, el tipo de la variable x se deduce como un número en el código siguiente.</span><span class="sxs-lookup"><span data-stu-id="26336-343">For example, the type of the x variable is inferred to be a number in the following code.</span></span>

```TypeScript
let x = 3;
//  ^ = let x: number
```

<span data-ttu-id="26336-344">Este tipo de inferencia tiene lugar al inicializar variables y miembros, establecer valores predeterminados de parámetro y determinar tipos de retorno de función.</span><span class="sxs-lookup"><span data-stu-id="26336-344">This kind of inference takes place when initializing variables and members, setting parameter default values, and determining function return types.</span></span>

### <a name="no-implicit-any-rule"></a><span data-ttu-id="26336-345">no-implicit-any rule</span><span class="sxs-lookup"><span data-stu-id="26336-345">no-implicit-any rule</span></span>

<span data-ttu-id="26336-346">Un script requiere los tipos de variables que se usan para declararse explícita o implícitamente.</span><span class="sxs-lookup"><span data-stu-id="26336-346">A script requires the types of the variables used to be explicitly or implicitly declared.</span></span> <span data-ttu-id="26336-347">Si el compilador de TypeScript no puede determinar el tipo de una variable (ya sea porque el tipo no se declara explícitamente o la inferencia de tipo no es posible), recibirá un error de tiempo de compilación (error antes de ejecutar el script).</span><span class="sxs-lookup"><span data-stu-id="26336-347">If the TypeScript compiler is unable to determine the type of a variable (either because type is not declared explicitly or type inference is not possible), then you will receive a compilation time error (error prior to running the script).</span></span> <span data-ttu-id="26336-348">También verá un error en el editor.</span><span class="sxs-lookup"><span data-stu-id="26336-348">You will see an error in the editor as well.</span></span>

:::image type="content" source="../../images/getting-started-iany.png" alt-text="Error implícito &quot;cualquiera&quot; que se muestra en el editor.":::

<span data-ttu-id="26336-350">Los scripts siguientes tienen errores de tiempo de compilación porque las variables se declaran sin tipos y TypeScript no puede determinar el tipo en el momento de la declaración.</span><span class="sxs-lookup"><span data-stu-id="26336-350">The following scripts have compilation time errors because variables are declared without types and TypeScript cannot determine the type at the time of declaration.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'value' gets 'any' type
    // because no type is declared.
    let value; 
    // Even when a number type is assigned,
    // the type of 'value' remains any.
    value = 10; 
    // The following statement fails because
    // Office Scripts can't send an argument
    // of type 'any' to Excel for processing.
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // The variable 'cell' gets 'any' type
    // because no type is defined.
    let cell; 
    cell = workbook.getActiveCell().getValue();
    // Office Scripts can't assign Range type object
    // to a variable of 'any' type.
    console.log(cell.getValue());
    return;
}
```

<span data-ttu-id="26336-351">Para evitar este error, use los siguientes patrones en su lugar.</span><span class="sxs-lookup"><span data-stu-id="26336-351">To avoid this error, use the following patterns instead.</span></span> <span data-ttu-id="26336-352">En cada caso, la variable y su tipo se declaran al mismo tiempo.</span><span class="sxs-lookup"><span data-stu-id="26336-352">In each case, the variable and its type are declared at the same time.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const value: number = 10; 
    workbook.getActiveCell().setValue(value);
    return;
}
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const cell: ExcelScript.Range = workbook.getActiveCell().getValue();
    console.log(cell.getValue()); 
    return;
}
```

## <a name="error-handling"></a><span data-ttu-id="26336-353">Control de errores</span><span class="sxs-lookup"><span data-stu-id="26336-353">Error handling</span></span>

<span data-ttu-id="26336-354">El error scripts de Office se puede clasificar en una de las siguientes categorías.</span><span class="sxs-lookup"><span data-stu-id="26336-354">Office Scripts error can be classified into one of the following categories.</span></span>

1. <span data-ttu-id="26336-355">Advertencia en tiempo de compilación que se muestra en el editor</span><span class="sxs-lookup"><span data-stu-id="26336-355">Compile-time warning shown in the editor</span></span>
1. <span data-ttu-id="26336-356">Error en tiempo de compilación que aparece cuando se ejecuta pero se produce antes de que comience la ejecución</span><span class="sxs-lookup"><span data-stu-id="26336-356">Compile-time error that appears when you run but occurs before execution begins</span></span>
1. <span data-ttu-id="26336-357">Error en tiempo de ejecución</span><span class="sxs-lookup"><span data-stu-id="26336-357">Runtime error</span></span>

<span data-ttu-id="26336-358">Las advertencias del editor se pueden identificar con los subrayados rojos ondulados del editor:</span><span class="sxs-lookup"><span data-stu-id="26336-358">Editor warnings can be identified using the wavy red underlines in the editor:</span></span>

:::image type="content" source="../../images/getting-started-eanyi.png" alt-text="Advertencia en tiempo de compilación que se muestra en el editor":::

<span data-ttu-id="26336-360">En ocasiones, también puede ver subrayados de advertencia naranja y mensajes informativos grises.</span><span class="sxs-lookup"><span data-stu-id="26336-360">At times, you may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="26336-361">Deben examinarse estrechamente aunque no van a causar errores.</span><span class="sxs-lookup"><span data-stu-id="26336-361">They should be examined closely though they are not going to cause errors.</span></span>

<span data-ttu-id="26336-362">No es posible distinguir entre errores en tiempo de compilación y en tiempo de ejecución, ya que ambos mensajes de error tienen un aspecto idéntico.</span><span class="sxs-lookup"><span data-stu-id="26336-362">It isn't possible to distinguish between compile-time and runtime errors as both error messages look identical.</span></span> <span data-ttu-id="26336-363">Ambos se producen cuando se ejecuta realmente el script.</span><span class="sxs-lookup"><span data-stu-id="26336-363">They both occur when you actually execute the script.</span></span> <span data-ttu-id="26336-364">En las imágenes siguientes se muestran ejemplos de un error en tiempo de compilación y un error en tiempo de ejecución.</span><span class="sxs-lookup"><span data-stu-id="26336-364">The following images show examples of a compile-time error and a runtime error.</span></span>

:::image type="content" source="../../images/getting-started-expany.png" alt-text="Ejemplo de error en tiempo de compilación":::

:::image type="content" source="../../images/getting-started-error-basic.png" alt-text="Ejemplo de error en tiempo de ejecución":::

<span data-ttu-id="26336-367">En ambos casos, verá el número de línea donde se produjo el error.</span><span class="sxs-lookup"><span data-stu-id="26336-367">In both cases, you will see the line number where the error occurred.</span></span> <span data-ttu-id="26336-368">A continuación, puede examinar el código, corregir el problema y volver a ejecutarlo.</span><span class="sxs-lookup"><span data-stu-id="26336-368">You can then examine the code, fix the issue, and run again.</span></span>

<span data-ttu-id="26336-369">Los siguientes son algunos procedimientos recomendados para evitar errores en tiempo de ejecución.</span><span class="sxs-lookup"><span data-stu-id="26336-369">Following are a few best practices to avoid runtime errors.</span></span>

### <a name="check-for-object-existence-before-deletion"></a><span data-ttu-id="26336-370">Comprobar la existencia de objetos antes de la eliminación</span><span class="sxs-lookup"><span data-stu-id="26336-370">Check for object existence before deletion</span></span>

<span data-ttu-id="26336-371">Como alternativa, para eliminar un objeto que puede existir o no, use este patrón:</span><span class="sxs-lookup"><span data-stu-id="26336-371">Alternatively, for deleting an object that may or may not exist, use this pattern:</span></span>

```TypeScript
// The ? ensures that the delete() API is only invoked if the object exists.
workbook.getWorksheet('Index')?.delete();

// Alternative:
const indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
    indexSheet.delete();
}
```

### <a name="do-pre-checks-at-the-beginning-of-the-script"></a><span data-ttu-id="26336-372">Realizar comprobaciones previas al principio del script</span><span class="sxs-lookup"><span data-stu-id="26336-372">Do pre-checks at the beginning of the script</span></span>

<span data-ttu-id="26336-373">Como práctica recomendada, asegúrese siempre de que todas las entradas estén presentes en el archivo de Excel antes de ejecutar el script.</span><span class="sxs-lookup"><span data-stu-id="26336-373">As a best practice, always ensure that all your inputs are present in the Excel file prior to running your script.</span></span> <span data-ttu-id="26336-374">Es posible que haya hecho ciertos supuestos acerca de los objetos que están presentes en el libro.</span><span class="sxs-lookup"><span data-stu-id="26336-374">You may have made certain assumptions about objects being present in the workbook.</span></span> <span data-ttu-id="26336-375">Si esos objetos no existen, es posible que el script encuentre un error al leer el objeto o sus datos.</span><span class="sxs-lookup"><span data-stu-id="26336-375">If those objects don't exist, your script may encounter an error when you read the object or its data.</span></span> <span data-ttu-id="26336-376">En lugar de iniciar el procesamiento y los errores en el medio después de que parte de las actualizaciones o procesamiento ya haya finalizado, es mejor realizar todas las comprobaciones previas al inicio del script.</span><span class="sxs-lookup"><span data-stu-id="26336-376">Rather than beginning the processing and erroring in the middle after part of the updates or processing has already finished, it is better to do all pre-checks at the start of the script.</span></span>

<span data-ttu-id="26336-377">Por ejemplo, el siguiente script requiere que se presenten dos tablas denominadas Table1 y Table2.</span><span class="sxs-lookup"><span data-stu-id="26336-377">For example, the following script requires two tables named Table1 and Table2 to be present.</span></span> <span data-ttu-id="26336-378">Por lo tanto, el script comprueba su presencia y termina con la instrucción y un `return` mensaje adecuado si no están presentes.</span><span class="sxs-lookup"><span data-stu-id="26336-378">Hence the script checks for their presence and ends with the `return` statement and an appropriate message if they are not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="26336-379">Si la comprobación para garantizar la presencia de datos de entrada se está produciendo en una función independiente, es importante finalizar el script emitiendo la `return` instrucción de la `main` función.</span><span class="sxs-lookup"><span data-stu-id="26336-379">If the verification to ensure the presence of input data is happening in a separate function, it's important to end the script by issuing the `return` statement from the `main` function.</span></span>

<span data-ttu-id="26336-380">En el ejemplo siguiente, la `main` función llama a la función para realizar las `inputPresent` comprobaciones previas.</span><span class="sxs-lookup"><span data-stu-id="26336-380">In the following example, the `main` function calls the `inputPresent` function to do the pre-checks.</span></span> <span data-ttu-id="26336-381">`inputPresent` devuelve un valor booleano ( o ) que indica `true` si todas las entradas necesarias están `false` presentes o no.</span><span class="sxs-lookup"><span data-stu-id="26336-381">`inputPresent` returns a boolean (`true` or `false`) indicating whether all required inputs are present or not.</span></span> <span data-ttu-id="26336-382">A continuación, es responsabilidad de la función emitir la instrucción (es decir, desde dentro de la función) para `main` `return` finalizar el script `main` inmediatamente.</span><span class="sxs-lookup"><span data-stu-id="26336-382">It's then the responsibility of the `main` function to issue the `return` statement (that is, from within the `main` function) to end the script immediately.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }
  return true;
}
```

### <a name="when-to-abort-throw-the-script"></a><span data-ttu-id="26336-383">Cuándo anular ( `throw` ) el script</span><span class="sxs-lookup"><span data-stu-id="26336-383">When to abort (`throw`) the script</span></span>  

<span data-ttu-id="26336-384">En su mayoría, no es necesario anular ( `throw` ) desde el script.</span><span class="sxs-lookup"><span data-stu-id="26336-384">For the most part, you don't need to abort (`throw`) from your script.</span></span> <span data-ttu-id="26336-385">Esto se debe a que el script normalmente informa al usuario de que el script no se pudo ejecutar debido a un problema.</span><span class="sxs-lookup"><span data-stu-id="26336-385">This is because the script's usually informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="26336-386">En la mayoría de los casos, basta con finalizar el script con un mensaje de error y una `return` instrucción de la `main` función.</span><span class="sxs-lookup"><span data-stu-id="26336-386">In most case, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="26336-387">Sin embargo, si el script se ejecuta como parte de Power Automate, es posible que desee anular el flujo si no se cumplen determinadas condiciones.</span><span class="sxs-lookup"><span data-stu-id="26336-387">However, if your script is running as part of Power Automate, you may want to abort the flow if certain conditions are not met.</span></span> <span data-ttu-id="26336-388">Por lo tanto, es importante no realizar un error, sino emitir una instrucción para anular el script para que no se ejecuten las instrucciones de `return` `throw` código posteriores.</span><span class="sxs-lookup"><span data-stu-id="26336-388">It's therefore important to not `return` upon an error but rather issue a `throw` statement to abort the script so that any subsequent code statements don't run.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Tables that should be in the workbook for the script to work:
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  if (!targetTable || !sourceTable) {
    // Abort script.
    throw `Required tables missing - Check that both source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

<span data-ttu-id="26336-389">Como se mencionó en la sección siguiente, otro escenario es cuando tiene varias funciones implicadas (llamadas `main` que llaman, etc.) lo que dificulta la `functionX` `functionY` propagación del error.</span><span class="sxs-lookup"><span data-stu-id="26336-389">As mentioned in the following section, another scenario is when you have several functions involved (`main` calls `functionX` which calls `functionY`, etc.) which makes it hard to propagate the error.</span></span> <span data-ttu-id="26336-390">Aborting/throwing from the nested function with a message may be easier than returning an error all the way up to `main` and returning from with an error `main` message.</span><span class="sxs-lookup"><span data-stu-id="26336-390">Aborting/throwing from the nested function with a message may be easier than returning an error all the way up to `main` and returning from `main` with an error message.</span></span>

### <a name="when-to-use-trycatch-throw-exception"></a><span data-ttu-id="26336-391">Cuándo usar try.. catch (excepción de lanzamiento)</span><span class="sxs-lookup"><span data-stu-id="26336-391">When to use try..catch (throw exception)</span></span>

<span data-ttu-id="26336-392">La técnica es una forma de detectar si se produjo un error en una llamada [`try..catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) API y controlar ese error en el script.</span><span class="sxs-lookup"><span data-stu-id="26336-392">The [`try..catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) technique is a way to detect if an API call failed and handle that error in your script.</span></span> <span data-ttu-id="26336-393">Puede ser importante comprobar el valor devuelto de una API para comprobar que se completó correctamente.</span><span class="sxs-lookup"><span data-stu-id="26336-393">It may be important to check the return value of an API to verify that it was completed successfully.</span></span>

<span data-ttu-id="26336-394">Tenga en cuenta el siguiente fragmento de código de ejemplo.</span><span class="sxs-lookup"><span data-stu-id="26336-394">Consider the following example snippet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Somewhere in the script, perform a large data update.
  range.setValues(someLargeValues);

}
```

<span data-ttu-id="26336-395">La `setValues()` llamada puede producir un error y provocar el error de script.</span><span class="sxs-lookup"><span data-stu-id="26336-395">The `setValues()` call may fail and result in the script failure.</span></span> <span data-ttu-id="26336-396">Es posible que desee controlar esta condición en el código y quizás personalizar el mensaje de error o dividir la actualización en unidades más pequeñas, etc. En ese caso, es importante saber que la API ha devuelto un error e interpretar o controlar ese error.</span><span class="sxs-lookup"><span data-stu-id="26336-396">You may wish to handle this condition in your code and perhaps customize the error message or break up the update into smaller units, etc. In that case, it's important to know that the API returned an error and interpret or handle that error.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Please inspect and run again.`);
    console.log(error);
    return; // End script (assuming this is in main function).
}

// OR...

try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ____. Trying a different approach`);
    handleUpdatesInSmallerChunks(someLargeValues);
}

// Continue...
}
```

<span data-ttu-id="26336-397">Otro escenario es cuando la función principal llama a otra función, que a su vez llama a otra función (y así sucesivamente). Y la llamada a la API que le interesa se produce en la función inferior.</span><span class="sxs-lookup"><span data-stu-id="26336-397">Another scenario is when main function calls another function, which in turn calls another function (and so on..), and the API call that you care about happens down in the bottom function.</span></span> <span data-ttu-id="26336-398">Propagar el error hasta el final `main` puede que no sea factible o conveniente.</span><span class="sxs-lookup"><span data-stu-id="26336-398">Propagating the error all the way up to `main` may not be feasible or convenient.</span></span> <span data-ttu-id="26336-399">En ese caso, es más conveniente lanzar un error en la función inferior.</span><span class="sxs-lookup"><span data-stu-id="26336-399">In that case, throwing an error in the bottom function will be most convenient.</span></span>

```TypeScript

function main(workbook: ExcelScript.Workbook) {
    ...
    updateRangeInChunks(sheet.getRange("B1"), data);
    ...
}

function updateRangeInChunks(
    ...
    updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
    ...
}

function updateTargetRange(
      targetCell: ExcelScript.Range,
      values: (string | boolean | number)[][]
    ) {
    const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
    console.log(`Updating the range: ${targetRange.getAddress()}`);
    try {
      targetRange.setValues(values);
    } catch (e) {
      throw `Error while updating the whole range: ${JSON.stringify(e)}`;
    }
    return;
}
```

<span data-ttu-id="26336-400">*Advertencia:* El `try..catch` uso dentro de un bucle ralentizará el script.</span><span class="sxs-lookup"><span data-stu-id="26336-400">*Warning*: Using `try..catch` inside of a loop will slow down your script.</span></span> <span data-ttu-id="26336-401">Evite usar esto dentro o alrededor de bucles.</span><span class="sxs-lookup"><span data-stu-id="26336-401">Avoid using this inside of or around loops.</span></span>

## <a name="range-basics"></a><span data-ttu-id="26336-402">Conceptos básicos sobre rangos</span><span class="sxs-lookup"><span data-stu-id="26336-402">Range basics</span></span>

<span data-ttu-id="26336-403">Consulte [Range Basics](range-basics.md) antes de ir más lejos en su viaje.</span><span class="sxs-lookup"><span data-stu-id="26336-403">Check out [Range Basics](range-basics.md) before you go further on your journey.</span></span>

## <a name="basic-performance-considerations"></a><span data-ttu-id="26336-404">Consideraciones básicas de rendimiento</span><span class="sxs-lookup"><span data-stu-id="26336-404">Basic performance considerations</span></span>

### <a name="avoid-slow-operations-in-the-loop"></a><span data-ttu-id="26336-405">Evitar operaciones lentas en el bucle</span><span class="sxs-lookup"><span data-stu-id="26336-405">Avoid slow operations in the loop</span></span>

<span data-ttu-id="26336-406">Ciertas operaciones cuando se realizan dentro o alrededor de las instrucciones de bucle, como , , , , etc. pueden provocar `for` `for..of` un rendimiento `map` `forEach` lento.</span><span class="sxs-lookup"><span data-stu-id="26336-406">Certain operations when done inside/around the loop statements such as `for`, `for..of`, `map`, `forEach`, etc. can lead to slow performance.</span></span> <span data-ttu-id="26336-407">Evite las siguientes categorías de API.</span><span class="sxs-lookup"><span data-stu-id="26336-407">Avoid the following API categories.</span></span>

* <span data-ttu-id="26336-408">`get*` API</span><span class="sxs-lookup"><span data-stu-id="26336-408">`get*` APIs</span></span>

<span data-ttu-id="26336-409">Lea todos los datos que necesite fuera del bucle en lugar de leerlo dentro del bucle.</span><span class="sxs-lookup"><span data-stu-id="26336-409">Read all the data you need outside of the loop rather than reading it inside of the loop.</span></span> <span data-ttu-id="26336-410">A veces, es difícil evitar leer dentro de bucles; en tal caso, asegúrese de que los recuentos de bucles no son demasiado grandes o de administrarlos en lotes para evitar tener que recorrer una estructura de datos grande.</span><span class="sxs-lookup"><span data-stu-id="26336-410">At times, it is hard to avoid reading inside of loops; in such a case, make sure your loop counts are not too large or manage them in batches to avoid having to loop through a large data structure.</span></span>

<span data-ttu-id="26336-411">**Nota:** Si el rango o los datos con los que está trabajando es bastante grande (por ejemplo, >celdas de 100K), es posible que necesite usar técnicas avanzadas como dividir las lecturas y escrituras en varios fragmentos.</span><span class="sxs-lookup"><span data-stu-id="26336-411">**Note**: If the range/data you are dealing with is quite large (say >100K cells), you may need to use advanced techniques like breaking up your read/writes into multiple chunks.</span></span> <span data-ttu-id="26336-412">El siguiente vídeo es realmente para una configuración de datos de tamaño pequeño y mediano.</span><span class="sxs-lookup"><span data-stu-id="26336-412">The following video is really for a small-mid sized data setup.</span></span> <span data-ttu-id="26336-413">Para un conjunto de datos de gran tamaño, consulte [escenario de escritura de datos avanzado.](write-large-dataset.md)</span><span class="sxs-lookup"><span data-stu-id="26336-413">For a large dataset, refer to [advanced data write scenario](write-large-dataset.md).</span></span>

<span data-ttu-id="26336-414">[![Vídeo que proporciona una sugerencia de optimización de lectura y escritura](../../images/getting-started-v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "Vídeo que muestra la sugerencia de optimización de lectura y escritura")</span><span class="sxs-lookup"><span data-stu-id="26336-414">[![Video providing a read-and-write optimization tip](../../images/getting-started-v_perf.jpg)](https://youtu.be/lsR_GvVW3Pg "Video showing read-and-write optimization tip")</span></span>

* <span data-ttu-id="26336-415">`console.log` instrucción (vea el ejemplo siguiente)</span><span class="sxs-lookup"><span data-stu-id="26336-415">`console.log` statement (see the following example)</span></span>

```TypeScript
// Color each cell with random color.
for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
        range
            .getCell(row, col)
            .getFormat()
            .getFill()
            .setColor(`#${Math.random().toString(16).substr(-6)}`);
        /* Avoid such console.log inside loop */
        // console.log("Updating" + range.getCell(row, col).getAddress());
    }
}
```

* <span data-ttu-id="26336-416">`try {} catch ()` instrucción</span><span class="sxs-lookup"><span data-stu-id="26336-416">`try {} catch ()` statement</span></span>

<span data-ttu-id="26336-417">Evite bucles de control `for` de excepciones.</span><span class="sxs-lookup"><span data-stu-id="26336-417">Avoid exception handling `for` loops.</span></span> <span data-ttu-id="26336-418">Bucles tanto interiores como externos.</span><span class="sxs-lookup"><span data-stu-id="26336-418">Both inside and outside loops.</span></span>

## <a name="note-to-vba-developers"></a><span data-ttu-id="26336-419">Nota para desarrolladores de VBA</span><span class="sxs-lookup"><span data-stu-id="26336-419">Note to VBA developers</span></span>

<span data-ttu-id="26336-420">El lenguaje TypeScript difiere de VBA tanto sintácticamente como en convenciones de nomenclatura.</span><span class="sxs-lookup"><span data-stu-id="26336-420">The TypeScript language differs from VBA both syntactically as well as in naming conventions.</span></span>

<span data-ttu-id="26336-421">Consulte los siguientes fragmentos de código equivalentes.</span><span class="sxs-lookup"><span data-stu-id="26336-421">Check out the following equivalent snippets.</span></span>

```vba
Worksheets("Sheet1").Range("A1:G37").Clear
```

```TypeScript
workbook.getWorksheet('Sheet1').getRange('A1:G37').clear(ExcelScript.ClearApplyTo.all);
```

<span data-ttu-id="26336-422">Algunas cosas que se pueden llamar acerca de TypeScript:</span><span class="sxs-lookup"><span data-stu-id="26336-422">A few things to call out about TypeScript:</span></span>

* <span data-ttu-id="26336-423">Es posible que observe que todos los métodos necesitan tener paréntesis de apertura y cierre para ejecutarse.</span><span class="sxs-lookup"><span data-stu-id="26336-423">You may notice that all methods need to have open-close parentheses to execute.</span></span> <span data-ttu-id="26336-424">Los argumentos se pasan de forma idéntica, pero es posible que algunos argumentos sean necesarios para la ejecución (es decir, obligatorios frente a opcionales).</span><span class="sxs-lookup"><span data-stu-id="26336-424">Arguments are passed identically but some arguments may be required for execution (that is, required vs optional).</span></span>
* <span data-ttu-id="26336-425">La convención de nomenclatura sigue a camelCase en lugar de a la convención de PascalCase.</span><span class="sxs-lookup"><span data-stu-id="26336-425">The naming convention follows camelCase instead of PascalCase convention.</span></span>
* <span data-ttu-id="26336-426">Los métodos suelen `get` tener `set` o prefijos que indican si está leyendo o escribiendo miembros de objeto.</span><span class="sxs-lookup"><span data-stu-id="26336-426">Methods usually have `get` or `set` prefixes indicating whether it is reading or writing object members.</span></span>
* <span data-ttu-id="26336-427">Los bloques de código se definen e identifican mediante llaves de cierre abierto: `{` `}` .</span><span class="sxs-lookup"><span data-stu-id="26336-427">The code blocks are defined and identified by open-close curly braces: `{` `}`.</span></span> <span data-ttu-id="26336-428">Los bloques son necesarios `if` para condiciones, `while` instrucciones, `for` bucles, definiciones de función, etc.</span><span class="sxs-lookup"><span data-stu-id="26336-428">Blocks are required for `if` conditions, `while` statements, `for` loops, function definitions, etc.</span></span>
* <span data-ttu-id="26336-429">Las funciones pueden llamar a otras funciones e incluso puede definir funciones dentro de una función.</span><span class="sxs-lookup"><span data-stu-id="26336-429">Functions can call other functions and you can even define functions within a function.</span></span>

<span data-ttu-id="26336-430">En general, TypeScript es un idioma diferente y hay pocas similitudes entre ellos.</span><span class="sxs-lookup"><span data-stu-id="26336-430">Overall, TypeScript is a different language and there are few similarities between them.</span></span> <span data-ttu-id="26336-431">Sin embargo, la PROPIA API de scripts de Office usa una terminología y una jerarquía de modelos de datos (modelo de objetos) similares a las API de VBA y eso le ayudará a navegar.</span><span class="sxs-lookup"><span data-stu-id="26336-431">However, the Office Scripts API themselves use similar terminology and data-model (object model) hierarchy as VBA APIs and that should help you navigate around.</span></span>
