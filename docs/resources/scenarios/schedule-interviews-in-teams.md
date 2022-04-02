---
title: Programar entrevistas en Teams
description: Obtenga información sobre cómo usar Office scripts para enviar una reunión Teams desde Excel datos.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 25b70f2ee3f71c101d4ee20068c020edb5e0ac77
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585432"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Office ejemplo scripts: Programar entrevistas en Teams

En este escenario, es un reclutador de recursos humanos que programa reuniones de entrevista con candidatos en Teams. La programación de entrevistas de los candidatos se administra en un Excel. Deberá enviar la invitación a la reunión Teams a los candidatos y entrevistadores. A continuación, debe actualizar el Excel con la confirmación de que Teams reuniones se han enviado.

La solución tiene tres pasos que se combinan en un solo Power Automate flujo.

1. Un script extrae datos de una tabla y devuelve una matriz de objetos como datos JSON.
1. A continuación, los datos se envían **al Teams Crear una Teams de reunión** para enviar invitaciones.
1. Los mismos datos JSON se envían a otro script para actualizar el estado de la invitación.

## <a name="scripting-skills-covered"></a>Habilidades de scripting cubiertas

* Power Automate flujos
* Teams integración
* Análisis de tablas

## <a name="sample-excel-file"></a>Archivo Excel ejemplo

Descarga el archivo <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> usado en esta solución y pruébalo tú mismo. Asegúrese de cambiar al menos una de las direcciones de correo electrónico para que reciba una invitación.

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>Código de ejemplo: extraer datos de tabla para programar invitaciones

Agregue este script a la colección de scripts. Así mismo, **asigne el nombre Programar entrevistas** para el flujo.

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

## <a name="sample-code-mark-rows-as-invited"></a>Código de ejemplo: Marcar filas como invitadas

Agregue este script a la colección de scripts. Así lo llama **Record Sent Invites** for the flow.

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>Flujo de ejemplo: ejecute los scripts de programación de entrevistas y envíe las Teams reuniones

1. Cree un nuevo **flujo de nube instantánea**.
1. Elija **Desencadenar manualmente un flujo y** seleccione **Crear**.
1. Agregue un **nuevo paso que** use el **conector Excel online (empresa)** y la **acción Ejecutar script**. Complete el conector con los siguientes valores.
    1. **Ubicación**: OneDrive para la Empresa
    1. **Biblioteca de documentos**: OneDrive
    1. **Archivo**: hr-interviews.xlsx *(elegido a través del explorador de archivos)*
    1. **Script**: Programar entrevistas Captura de pantalla del conector Excel Online (Empresa) para obtener datos de entrevista :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="del libro en Power Automate.":::
1. Agregue un **paso Nuevo que** use la **acción Crear una Teams reunión**. A medida que seleccione contenido dinámico en el conector Excel, se generará un valor **Aplicar a cada** bloque para el flujo. Complete el conector con los siguientes valores.
    1. **Identificador de calendario**: Calendario
    1. **Asunto**: Entrevista de Contoso
    1. **Mensaje**: **Mensaje** (el Excel valor)
    1. **Zona horaria**: hora estándar del Pacífico
    1. **Hora de** inicio: **StartTime** (el Excel valor)
    1. **Hora de finalización**: **FinishTime** (el Excel valor)
    1. **Asistentes necesarios**: **CandidateEmail** ; **InterviewerEmail** (los valores Excel) Captura de pantalla del conector Teams para programar reuniones :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="en Power Automate.":::
1. En el mismo **Aplicar a cada** bloque, agregue otro **conector Excel Online (Empresa)** con la **acción Ejecutar script**. Use los siguientes valores.
    1. **Ubicación**: OneDrive para la Empresa
    1. **Biblioteca de documentos**: OneDrive
    1. **Archivo**: hr-interviews.xlsx *(elegido a través del explorador de archivos)*
    1. **Script**: Registrar invitaciones enviadas
    1. **invites**: **resultado** (el valor Excel) Captura de pantalla del conector :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Excel Online (Empresa)"::: completado para registrar que las invitaciones se han enviado en Power Automate.
1. Guarde el flujo y pruébalo. Use el **botón Probar** en la página del editor de flujo o ejecute el flujo a través de la **pestaña Mis flujos** . Asegúrese de permitir el acceso cuando se le pida.

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Vídeo de aprendizaje: Enviar una reunión Teams desde Excel datos

[Vea a Sudhi Ramamurthy recorrer una versión de este ejemplo en YouTube](https://youtu.be/HyBdx52NOE8). Su versión usa un script más sólido que controla el cambio de columnas y los tiempos de reunión obsoletos.
