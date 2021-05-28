---
title: Programar entrevistas en Teams
description: Obtenga información sobre cómo usar Office scripts para enviar una reunión Teams desde Excel datos.
ms.date: 05/25/2021
localization_priority: Normal
ms.openlocfilehash: f93d9ceca6603ddb9e7123a393787fcf54597cca
ms.sourcegitcommit: 339ecbb9914d54f919e3475018888fb5d00abe89
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/28/2021
ms.locfileid: "52697800"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a><span data-ttu-id="2445d-103">Office Escenario de ejemplo scripts: Programar entrevistas en Teams</span><span class="sxs-lookup"><span data-stu-id="2445d-103">Office Scripts sample scenario: Schedule interviews in Teams</span></span>

<span data-ttu-id="2445d-104">En este escenario, es un reclutador de recursos humanos que programa reuniones de entrevista con candidatos en Teams.</span><span class="sxs-lookup"><span data-stu-id="2445d-104">In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams.</span></span> <span data-ttu-id="2445d-105">La programación de entrevistas de los candidatos se administra en un Excel.</span><span class="sxs-lookup"><span data-stu-id="2445d-105">You manage the interview schedule of candidates in an Excel file.</span></span> <span data-ttu-id="2445d-106">Deberá enviar la invitación a la reunión Teams al candidato y a los entrevistadores.</span><span class="sxs-lookup"><span data-stu-id="2445d-106">You'll need to send the Teams meeting invite to both the candidate and interviewers.</span></span> <span data-ttu-id="2445d-107">A continuación, debe actualizar el Excel con la confirmación de que Teams reuniones se han enviado.</span><span class="sxs-lookup"><span data-stu-id="2445d-107">You then need to update the Excel file with the confirmation that Teams meetings have been sent.</span></span>

<span data-ttu-id="2445d-108">La solución tiene tres pasos que se combinan en un solo Power Automate flujo.</span><span class="sxs-lookup"><span data-stu-id="2445d-108">The solution has three steps that are combined in a single Power Automate flow.</span></span>

1. <span data-ttu-id="2445d-109">Un script extrae datos de una tabla y devuelve una matriz de objetos como datos JSON.</span><span class="sxs-lookup"><span data-stu-id="2445d-109">A script extracts data from a table and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="2445d-110">A continuación, los datos se envían al Teams **Crear una Teams de reunión** para enviar invitaciones.</span><span class="sxs-lookup"><span data-stu-id="2445d-110">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span>
1. <span data-ttu-id="2445d-111">Los mismos datos JSON se envían a otro script para actualizar el estado de la invitación.</span><span class="sxs-lookup"><span data-stu-id="2445d-111">The same JSON data is sent to another script to update the status of the invitation.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="2445d-112">Habilidades de scripting cubiertas</span><span class="sxs-lookup"><span data-stu-id="2445d-112">Scripting skills covered</span></span>

* <span data-ttu-id="2445d-113">Power Automate flujos</span><span class="sxs-lookup"><span data-stu-id="2445d-113">Power Automate flows</span></span>
* <span data-ttu-id="2445d-114">Teams integración</span><span class="sxs-lookup"><span data-stu-id="2445d-114">Teams integration</span></span>
* <span data-ttu-id="2445d-115">Análisis de tablas</span><span class="sxs-lookup"><span data-stu-id="2445d-115">Table parsing</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="2445d-116">Archivo Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="2445d-116">Sample Excel file</span></span>

<span data-ttu-id="2445d-117">Descargue el archivo <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> se usa en esta solución y pruébalo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="2445d-117">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span> <span data-ttu-id="2445d-118">Asegúrese de cambiar al menos una de las direcciones de correo electrónico para que reciba una invitación.</span><span class="sxs-lookup"><span data-stu-id="2445d-118">Be sure to change at least one of the email addresses so that you receive an invite.</span></span>

## <a name="sample-code-extract-table-data-to-schedule-invites"></a><span data-ttu-id="2445d-119">Código de ejemplo: extraer datos de tabla para programar invitaciones</span><span class="sxs-lookup"><span data-stu-id="2445d-119">Sample code: Extract table data to schedule invites</span></span>

<span data-ttu-id="2445d-120">Asigne a este script **el nombre Programar entrevistas** para el flujo.</span><span class="sxs-lookup"><span data-stu-id="2445d-120">Name this script **Schedule Interviews** for the flow.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  const MEETING_DURATION = workbook.getWorksheet("Constants").getRange("B1").getValue() as number;
  const MESSAGE_TEMPLATE = workbook.getWorksheet("Constants").getRange("B2").getValue() as string;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet("Interviews");
  const table = sheet.getTables()[0];
  const dataRows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Convert the table rows into InterviewInvite objects for the flow.
  let invites: InterviewInvite[] = [];
  dataRows.forEach((row) => {
    const inviteSent = row[1] as boolean;
    if (!inviteSent) {
      const startTime = new Date(Math.round(((row[6] as number) - 25569) * 86400 * 1000));
      const finishTime = new Date(startTime.getTime() + MEETING_DURATION * 60 * 1000);
      const candidateName = row[2] as string;
      const interviewerName = row[4] as string;

      invites.push({
        ID: row[0] as string,
        Candidate: candidateName,
        CandidateEmail: row[3] as string,
        Interviewer: row[4] as string,
        InterviewerEmail: row[5] as string,
        StartTime: startTime.toISOString(),
        FinishTime: finishTime.toISOString(),
        Message: generateInviteMessage(MESSAGE_TEMPLATE, candidateName, interviewerName)
      });
    }    
  });

  console.log(JSON.stringify(invites));
  return invites;
}

function generateInviteMessage(
  messageTemplate: string,
   candidate: string,
   interviewer: string) : string {
  return messageTemplate.replace("_Candidate_", candidate).replace("_Interviewer_", interviewer);
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-code-mark-rows-as-invited"></a><span data-ttu-id="2445d-121">Código de ejemplo: Marcar filas como invitadas</span><span class="sxs-lookup"><span data-stu-id="2445d-121">Sample code: Mark rows as invited</span></span>

<span data-ttu-id="2445d-122">Asigne a este script **el nombre Record Sent Invites** para el flujo.</span><span class="sxs-lookup"><span data-stu-id="2445d-122">Name this script **Record Sent Invites** for the flow.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, invites: InterviewInvite[]) {
  const table = workbook.getWorksheet("Interviews").getTables()[0];

  // Get the ID and Invite Sent columns from the table.
  const idColumn = table.getColumnByName("ID");
  const idRange = idColumn.getRangeBetweenHeaderAndTotal().getValues();
  const inviteSentColumn = table.getColumnByName("Invite Sent?");

  const dataRowCount = idRange.length;

  // Find matching IDs to mark the correct row.
  for (let row = 0; row < dataRowCount; row++){
    let inviteSent = invites.find((invite) => {
      return invite.ID == idRange[row][0] as string;
    });

    if (inviteSent) {
      inviteSentColumn.getRangeBetweenHeaderAndTotal().getCell(row, 0).setValue(true);
      console.log(`Invite for ${inviteSent.Candidate} has been sent.`);
    }
  } 
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a><span data-ttu-id="2445d-123">Flujo de ejemplo: ejecute los scripts de programación de entrevistas y envíe las Teams reuniones</span><span class="sxs-lookup"><span data-stu-id="2445d-123">Sample flow: Run the interview scheduling scripts and send the Teams meetings</span></span>

1. <span data-ttu-id="2445d-124">Crear un nuevo **flujo de nube instantánea**.</span><span class="sxs-lookup"><span data-stu-id="2445d-124">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="2445d-125">Seleccione **Desencadenar manualmente un flujo y** presione **Crear**.</span><span class="sxs-lookup"><span data-stu-id="2445d-125">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="2445d-126">Agregue un **paso Nuevo que** use el conector Excel online **(empresa)** y la **acción Ejecutar script.**</span><span class="sxs-lookup"><span data-stu-id="2445d-126">Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action.</span></span> <span data-ttu-id="2445d-127">Complete el conector con los siguientes valores.</span><span class="sxs-lookup"><span data-stu-id="2445d-127">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="2445d-128">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="2445d-128">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="2445d-129">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="2445d-129">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="2445d-130">**Archivo**: hr-interviews.xlsx *(elegido a través del explorador de archivos)*</span><span class="sxs-lookup"><span data-stu-id="2445d-130">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. **Script:** Programar entrevistas Captura de pantalla del conector :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Excel Online (Empresa)"::: completado para obtener datos de entrevista del libro en Power Automate
1. <span data-ttu-id="2445d-132">Agregue un **paso Nuevo** que use la acción Crear una **Teams reunión.**</span><span class="sxs-lookup"><span data-stu-id="2445d-132">Add a **New step** that uses the **Create a Teams meeting** action.</span></span> <span data-ttu-id="2445d-133">A medida que seleccione contenido dinámico en el conector Excel, se generará un valor Aplicar a **cada** bloque para el flujo.</span><span class="sxs-lookup"><span data-stu-id="2445d-133">As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow.</span></span> <span data-ttu-id="2445d-134">Complete el conector con los siguientes valores.</span><span class="sxs-lookup"><span data-stu-id="2445d-134">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="2445d-135">**Identificador de calendario**: Calendario</span><span class="sxs-lookup"><span data-stu-id="2445d-135">**Calendar id**: Calendar</span></span>
    1. <span data-ttu-id="2445d-136">**Asunto**: Entrevista de Contoso</span><span class="sxs-lookup"><span data-stu-id="2445d-136">**Subject**: Contoso Interview</span></span>
    1. <span data-ttu-id="2445d-137">**Message**: **Message** (el Excel valor)</span><span class="sxs-lookup"><span data-stu-id="2445d-137">**Message**: **Message** (the Excel value)</span></span>
    1. <span data-ttu-id="2445d-138">**Zona horaria:** hora estándar del Pacífico</span><span class="sxs-lookup"><span data-stu-id="2445d-138">**Time zone**: Pacific Standard Time</span></span>
    1. <span data-ttu-id="2445d-139">**Hora de** inicio: **StartTime** (el Excel valor)</span><span class="sxs-lookup"><span data-stu-id="2445d-139">**Start time**: **StartTime** (the Excel value)</span></span>
    1. <span data-ttu-id="2445d-140">**Hora de** finalización: **FinishTime** (el Excel valor)</span><span class="sxs-lookup"><span data-stu-id="2445d-140">**End time**: **FinishTime** (the Excel value)</span></span>
    1. **Asistentes requeridos**: **CandidateEmail** ; **InterviewerEmail** (los valores Excel) Captura de pantalla del conector de Teams para programar :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="reuniones en Power Automate":::
1. <span data-ttu-id="2445d-142">En el mismo **Aplicar a cada** bloque, agregue otro conector Excel Online **(Empresa)** con la **acción Ejecutar script.**</span><span class="sxs-lookup"><span data-stu-id="2445d-142">In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="2445d-143">Use los siguientes valores.</span><span class="sxs-lookup"><span data-stu-id="2445d-143">Use the following values.</span></span>
    1. <span data-ttu-id="2445d-144">**Ubicación**: OneDrive para la Empresa</span><span class="sxs-lookup"><span data-stu-id="2445d-144">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="2445d-145">**Biblioteca de documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="2445d-145">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="2445d-146">**Archivo**: hr-interviews.xlsx *(elegido a través del explorador de archivos)*</span><span class="sxs-lookup"><span data-stu-id="2445d-146">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. <span data-ttu-id="2445d-147">**Script**: Registrar invitaciones enviadas</span><span class="sxs-lookup"><span data-stu-id="2445d-147">**Script**: Record Sent Invites</span></span>
    1. **invites**: **result** (el valor Excel) Captura de pantalla del conector :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Excel Online (Empresa)"::: completado para registrar que las invitaciones se han enviado en Power Automate
1. <span data-ttu-id="2445d-149">Guarde el flujo y pruébalo.</span><span class="sxs-lookup"><span data-stu-id="2445d-149">Save the flow and try it out.</span></span>

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="2445d-150">Vídeo de aprendizaje: Enviar una reunión Teams desde Excel datos</span><span class="sxs-lookup"><span data-stu-id="2445d-150">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="2445d-151">[Vea el recorrido de Sudhi Ramamurthy a través](https://youtu.be/HyBdx52NOE8)de una versión de este ejemplo en YouTube .</span><span class="sxs-lookup"><span data-stu-id="2445d-151">[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span> <span data-ttu-id="2445d-152">Su versión usa un script más sólido que controla el cambio de columnas y los tiempos de reunión obsoletos.</span><span class="sxs-lookup"><span data-stu-id="2445d-152">His version uses a more robust script that handles changing columns and obsolete meeting times.</span></span>
