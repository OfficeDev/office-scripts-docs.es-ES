---
title: Enviar una reunión Teams desde Excel datos
description: Obtenga información sobre cómo usar Office scripts para enviar una reunión Teams desde Excel datos.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b0a3d5732727fd399fe34f3645336840ba4c156d
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232735"
---
# <a name="send-teams-meeting-from-excel-data"></a><span data-ttu-id="8450b-103">Enviar Teams reunión desde Excel datos</span><span class="sxs-lookup"><span data-stu-id="8450b-103">Send Teams meeting from Excel data</span></span>

<span data-ttu-id="8450b-104">Esta solución muestra cómo usar Office scripts y acciones Power Automate para seleccionar filas de un archivo Excel y usarlo para enviar una invitación Teams reunión y, a continuación, actualizar Excel.</span><span class="sxs-lookup"><span data-stu-id="8450b-104">This solution shows how to use Office Scripts and Power Automate actions to select rows from Excel file and use it to send a Teams meeting invite then update Excel.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="8450b-105">Ejemplo ficticio</span><span class="sxs-lookup"><span data-stu-id="8450b-105">Example scenario</span></span>

* <span data-ttu-id="8450b-106">Un reclutador de recursos humanos administra la programación de entrevistas de los candidatos en un Excel.</span><span class="sxs-lookup"><span data-stu-id="8450b-106">An HR recruiter manages the interview schedule of candidates in an Excel file.</span></span>
* <span data-ttu-id="8450b-107">El reclutador debe enviar la invitación Teams reunión al candidato y a los entrevistadores.</span><span class="sxs-lookup"><span data-stu-id="8450b-107">The recruiter needs to send the Teams meeting invite to the candidate and interviewers.</span></span> <span data-ttu-id="8450b-108">Las reglas de negocio son seleccionar:</span><span class="sxs-lookup"><span data-stu-id="8450b-108">The business rules are to select:</span></span>

    <span data-ttu-id="8450b-109">(a) Invita solo a aquellos para los que la invitación no se haya enviado ya como se registra en la columna de archivo.</span><span class="sxs-lookup"><span data-stu-id="8450b-109">(a) Invites to only those for whom the invite isn't already sent as recorded in the file column.</span></span>

    <span data-ttu-id="8450b-110">(b) Fechas de entrevista en el futuro (sin fechas anteriores).</span><span class="sxs-lookup"><span data-stu-id="8450b-110">(b) Interview dates in the future (no past dates).</span></span>

* <span data-ttu-id="8450b-111">El reclutador debe actualizar el archivo Excel con la confirmación de que todas Teams reuniones se han enviado para los registros elegibles.</span><span class="sxs-lookup"><span data-stu-id="8450b-111">The recruiter needs to update the Excel file with the confirmation that all Teams meetings have been sent for the eligible records.</span></span>

<span data-ttu-id="8450b-112">La solución tiene 3 partes:</span><span class="sxs-lookup"><span data-stu-id="8450b-112">The solution has 3 parts:</span></span>

1. <span data-ttu-id="8450b-113">Office Script para extraer datos de una tabla en función de las condiciones y devuelve una matriz de objetos como datos JSON.</span><span class="sxs-lookup"><span data-stu-id="8450b-113">Office Script to extract data from a table based on conditions and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="8450b-114">A continuación, los datos se envían al Teams **Crear una Teams de reunión** para enviar invitaciones.</span><span class="sxs-lookup"><span data-stu-id="8450b-114">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span> <span data-ttu-id="8450b-115">Envíe una Teams reunión por instancia en la matriz JSON.</span><span class="sxs-lookup"><span data-stu-id="8450b-115">Send one Teams meeting per instance in the JSON array.</span></span>
1. <span data-ttu-id="8450b-116">Envíe los mismos datos JSON a otro script Office para actualizar el estado de la invitación.</span><span class="sxs-lookup"><span data-stu-id="8450b-116">Send the same JSON data to another Office Script to update the status of the invitation.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="8450b-117">Archivo Excel ejemplo</span><span class="sxs-lookup"><span data-stu-id="8450b-117">Sample Excel file</span></span>

<span data-ttu-id="8450b-118">Descargue el archivo <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> se usa en esta solución y pruébalo usted mismo.</span><span class="sxs-lookup"><span data-stu-id="8450b-118">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span>

## <a name="sample-code-select-filtered-rows-from-table-as-json"></a><span data-ttu-id="8450b-119">Código de ejemplo: seleccionar filas filtradas de la tabla como JSON</span><span class="sxs-lookup"><span data-stu-id="8450b-119">Sample code: Select filtered rows from table as JSON</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  console.log("Current date time: " + new Date().toUTCString())
  const MEETING_DURATION = workbook.getNamedItem('MeetingDuration').getRange().getValue() as number;
  const sheet = workbook.getWorksheet('Interviews');
  const table = sheet.getTables()[0];
  const dataRows: string[][] = table.getRange().getTexts();
  // OR use the following statement if there's no table:
  // let dataRows = sheet.getUsedRange().getValues();
  const selectedRows = dataRows.filter((row, i) => {
    // Select header row and any data row with the status column equal to approach value.
    return (row[1] === 'FALSE' || i === 0)
  })
  const recordDetails: RecordDetail[] = returnObjectFromValues(selectedRows as string[][]);
  const inviteRecords = generateInterviewRecords(recordDetails, MEETING_DURATION);
  console.log(JSON.stringify(inviteRecords));
  return inviteRecords;
}

/**
 * This helper function converts table values into an object array.
 */
function returnObjectFromValues(values: string[][]): RecordDetail[] {
  let objArray: BasicObj[] = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray as RecordDetail[];
}

/**
 * Generate interview records by selecting required columns.
 * @param records Input records
 * @param mins Number of minutes to add to the start date-time
 */
function generateInterviewRecords(records: RecordDetail[], mins: number): InterviewInvite[] {
  const interviewInvites: InterviewInvite[] = []

  records.forEach((record) => {
    // Interviewer 1
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time1'])) > new Date()) {
      console.log("selected " + new Date(record['Start time1']).toUTCString());
      let startTime = new Date(record['Start time1']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time1']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer1,
        InterviewerEmail: record['Interviewer1 email'],
        StartTime: startTime,
        FinishTime: finishTime
      })
    } else {
      console.log("Rejected " + (new Date(record['Start time1']).toUTCString()))
    }
    // Interviewer 2 
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time2'])) > new Date()) {
      console.log("selected " + new Date(record['Start time2']).toUTCString());


      let startTime = new Date(record['Start time2']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time2']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer2,
        InterviewerEmail: record['Interviewer2 email'],
        StartTime: startTime,
        FinishTime: finishTime
      })
    } else {
      console.log("Rejected " + (new Date(record['Start time2']).toUTCString()))

    }
  })
  return interviewInvites;
}

/**
 * Add minutes to start date-time.
 * @param startDateTime Start date-time
 * @param mins Minutes to add to the start date-time
 */
function addMins(startDateTime: Date, mins: number) {
  return new Date(startDateTime.getTime() + mins * 60 * 1000);
}

// Basic key-value pair object.
interface BasicObj {
  [key: string]: string | number | boolean
}

// Input record that matches the table data.
interface RecordDetail extends BasicObj {
  ID: string
  'Invite to interview': string
  Candidate: string
  'Candidate email': string
  'Candidate contact': string
  Interviewer1: string
  'Interviewer1 email': string
  Interviewer2: string
  'Interviewer2 email': string
  'Start time1': string
  'Start time2': string
}

// Output record.
interface InterviewInvite extends BasicObj {
  ID: string
  Candidate: string
  CandidateEmail: string
  CandidateContact: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
}
```

## <a name="sample-code-mark-as-invited"></a><span data-ttu-id="8450b-120">Código de ejemplo: Marcar como invitado</span><span class="sxs-lookup"><span data-stu-id="8450b-120">Sample code: Mark as invited</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, completedInvitesString: string) {
    completedInvitesString = `[
  {
    "ID": "10",
    "Candidate": "Adele ",
    "CandidateEmail": "AdeleV@M365x904181.OnMicrosoft.com",
    "CandidateContact": "1234567899",
    "Interviewer": "Megan",
    "InterviewerEmail": "MeganB@M365x904181.OnMicrosoft.com",
    "StartTime": "2020-11-03T18:30:00Z",
    "FinishTime": "2020-11-03T22:45:00Z"
  },
  {
    "ID": "30",
    "Candidate": "Allan ",
    "CandidateEmail": "AllanD@M365x904181.OnMicrosoft.com",
    "CandidateContact": "1234567978",
    "Interviewer": "Raul",
    "InterviewerEmail": "RaulR@M365x904181.OnMicrosoft.com",
    "StartTime": "2020-11-03T23:00:00Z",
    "FinishTime": "2020-11-03T23:45:00Z"
  }
]`;
    let completedInvites = JSON.parse(completedInvitesString) as InterviewInvite[];
    const sheet = workbook.getWorksheet('Interviews');
    const range = sheet.getTables()[0].getRange();
    const dataRows = range.getValues();
    for (let i=0; i < dataRows.length; i++) {
        for (let invite of completedInvites) {
            if (String(dataRows[i][0]) === invite.ID) {
                range.getCell(i,1).setValue(true);
            }
        }
    }
    return;
}


// Invite record.
interface InterviewInvite  {
    ID: string
    Candidate: string
    CandidateEmail: string
    CandidateContact: string
    Interviewer: string
    InterviewerEmail: string
    StartTime: string
    FinishTime: string
}
```

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="8450b-121">Vídeo de aprendizaje: Enviar una reunión Teams desde Excel datos</span><span class="sxs-lookup"><span data-stu-id="8450b-121">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="8450b-122">[Vea el recorrido de Sudhi Ramamurthy por este ejemplo en YouTube](https://youtu.be/HyBdx52NOE8).</span><span class="sxs-lookup"><span data-stu-id="8450b-122">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span>
