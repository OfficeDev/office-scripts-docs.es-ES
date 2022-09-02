---
title: Programar entrevistas en Teams
description: Obtenga información sobre cómo usar scripts de Office para enviar una reunión de Teams desde datos de Excel.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e8c4af40398842e219dc3e2a80c6d2ee72d6b83
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572580"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Escenario de ejemplo de Scripts de Office: Programar entrevistas en Teams

En este escenario, usted es un reclutador de RR. HH. que programa reuniones de entrevistas con candidatos en Teams. Administra la programación de entrevistas de los candidatos en un archivo de Excel. Tendrá que enviar la invitación a la reunión de Teams tanto a los candidatos como a los entrevistadores. A continuación, debe actualizar el archivo de Excel con la confirmación de que se han enviado reuniones de Teams.

La solución tiene tres pasos que se combinan en un único flujo de Power Automate.

1. Un script extrae datos de una tabla y devuelve una matriz de objetos como datos [JSON](https://www.w3schools.com/whatis/whatis_json.asp) .
1. A continuación, los datos se envían a la acción **Crear una reunión de Teams** para enviar invitaciones.
1. Los mismos datos JSON se envían a otro script para actualizar el estado de la invitación.

Para obtener más información sobre cómo trabajar con JSON, lea [Uso de JSON para pasar datos hacia y desde scripts de Office](../../develop/use-json.md).

## <a name="scripting-skills-covered"></a>Aptitudes de scripting cubiertas

* Flujos de Power Automate
* Integración de Teams
* Análisis de tablas

## <a name="sample-excel-file"></a>Archivo de Excel de ejemplo

Descargue el archivo [hr-schedule.xlsx](hr-schedule.xlsx) usado en esta solución y pruébelo usted mismo! Asegúrese de cambiar al menos una de las direcciones de correo electrónico para que reciba una invitación.

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>Código de ejemplo: Extracción de datos de tabla para programar invitaciones

Agregue este script a la colección de scripts. Asígnele el nombre **Programar entrevistas** para el flujo.

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

Agregue este script a la colección de scripts. Asígnele el nombre **Grabar invitaciones enviadas** para el flujo.

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>Flujo de ejemplo: ejecución de los scripts de programación de entrevistas y envío de las reuniones de Teams

1. Cree un flujo **de nube instantáneo**.
1. Elija **Desencadenar manualmente un flujo** y seleccione **Crear**.
1. Agregue un **nuevo paso** que use el conector **de Excel Online (Empresa)** y la acción **Ejecutar script** . Complete el conector con los siguientes valores.
    1. **Ubicación**: OneDrive para la Empresa
    1. **Biblioteca de documentos**: OneDrive
    1. **Archivo**: hr-interviews.xlsx *(elegido mediante el explorador de archivos)*
    1. **Script**: Programar entrevistas :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Captura de pantalla del conector de Excel Online (Empresa) completado para obtener datos de entrevistas del libro en Power Automate.":::
1. Agregue un **paso Nuevo** que use la acción **Crear una reunión de Teams** . Al seleccionar contenido dinámico desde el conector de Excel, se generará un objeto **Apply to each block (Aplicar a cada** bloque) para el flujo. Complete el conector con los siguientes valores.
    1. **Id. de calendario**: Calendario
    1. **Asunto**: Entrevista de Contoso
    1. **Mensaje**: **Mensaje** (el valor de Excel)
    1. **Zona horaria**: Hora estándar del Pacífico
    1. **Hora de inicio**: **StartTime** (el valor de Excel)
    1. **Hora de finalización**: **FinishTime** (el valor de Excel)
    1. **Asistentes necesarios**: **CandidateEmail** ; **InterviewerEmail** (los valores de Excel) :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="Captura de pantalla del conector de Teams completado para programar reuniones en Power Automate.":::
1. En el mismo bloque **Aplicar a cada** bloque, agregue otro conector **de Excel Online (Empresa)** con la acción **Ejecutar script** . Use los siguientes valores.
    1. **Ubicación**: OneDrive para la Empresa
    1. **Biblioteca de documentos**: OneDrive
    1. **Archivo**: hr-interviews.xlsx *(elegido mediante el explorador de archivos)*
    1. **Script**: invitaciones enviadas de registros
    1. **invite**: **resultado** (el valor de Excel) :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Captura de pantalla del conector de Excel Online (empresa) completado para registrar que las invitaciones se han enviado en Power Automate.":::
1. Guarde el flujo y pruébelo. Use el botón **Probar** de la página del editor de flujo o ejecute el flujo a través de la pestaña **Mis flujos** . Asegúrese de permitir el acceso cuando se le solicite.

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Vídeo de entrenamiento: Envío de una reunión de Teams desde datos de Excel

[Vea cómo Sudhi Ramamurthy recorre una versión de este ejemplo en YouTube](https://youtu.be/HyBdx52NOE8). Su versión usa un script más sólido que controla las columnas cambiantes y los tiempos de reunión obsoletos.
