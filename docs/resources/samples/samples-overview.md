---
title: Ejemplos de scripts de Office
description: Escenarios y ejemplos de Scripts de Office disponibles.
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5798da37bd4166d18b41c005c4d8cc8a4b6c401d
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572489"
---
# <a name="office-scripts-samples-and-scenarios"></a>Ejemplos y escenarios de Scripts de Office

Esta sección contiene soluciones de automatización [basadas en scripts de Office](../../overview/excel.md) que ayudan a los usuarios finales a lograr la automatización de las tareas diarias. Contiene escenarios realistas a los que se enfrentan los usuarios empresariales y proporciona soluciones detalladas junto con vínculos de vídeo de instrucciones paso a paso.

Para cada uno de los proyectos de [Conceptos básicos](#basics) y [Más allá de los conceptos básicos](#beyond-the-basics), consulte el código fuente, [**los vídeos de YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0) paso a paso y mucho más.

En Escenarios, hemos incluido algunos [ejemplos](#scenarios) de escenarios más grandes que muestran casos de uso reales.

También agradecemos [las contribuciones de la comunidad](#community-contributions-and-fun-samples). Estos ejemplos se código abierto.

> [!IMPORTANT]
> Asegúrese de cumplir los requisitos previos para los scripts de Office antes de probar los ejemplos. Los requisitos de la suscripción y la cuenta de Microsoft 365 se encuentran en la [sección "Requisitos" de información general de Office Scripts for Excel](../../overview/excel.md#requirements).

## <a name="basics"></a>Conceptos básicos

| Project | Detalles |
|---------|---------|
| [Conceptos básicos de scripting](excel-samples.md) | En estos ejemplos se muestran los bloques de creación fundamentales para los scripts de Office. |
| [Agregar comentarios en Excel](add-excel-comments.md) | En este ejemplo se agregan comentarios a una celda, incluido @mentioning un compañero. |
| [Agregar imágenes a un libro](add-image-to-workbook.md) | En este ejemplo se agrega una imagen a un libro y se copia una imagen entre hojas.|
| [Copia de varias tablas de Excel en una sola tabla](copy-tables-combine.md) | Este ejemplo combina datos de varias tablas de Excel en una sola tabla que incluye todas las filas. |
| [Crear una tabla de contenido de libro](table-of-contents.md) | En este ejemplo se crea una tabla de contenido con vínculos a cada hoja de cálculo. |
| [Quitar filtros de columna de tabla](clear-table-filter-for-active-cell.md) | En este ejemplo se borran todos los filtros de una columna de tabla. |
| [Registrar los cambios diarios en Excel y notificarlos con un flujo de Power Automate](report-day-to-day-changes.md) | En este ejemplo se usa un flujo programado de Power Automate para registrar lecturas diarias e informar de los cambios. |

## <a name="beyond-the-basics"></a>Más allá de los aspectos básicos

Consulte el siguiente proyecto de un extremo a otro que automatiza escenarios de ejemplo junto con scripts completos, archivos de Excel de ejemplo usados y [vídeos (hospedados en YouTube).](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Detalles |
|---------|---------|
| [Combinar hojas de cálculo en un solo libro](combine-worksheets-into-single-workbook.md) | En este ejemplo se usan scripts de Office y Power Automate para extraer datos de otros libros en un solo libro. |
| [Conversión de archivos CSV en libros de Excel](convert-csv.md) | En este ejemplo se usan scripts de Office y Power Automate para crear archivos .xlsx a partir de archivos .csv. |
| [Libros de referencia cruzada](excel-cross-reference.md) | En este ejemplo se usan scripts de Office y Power Automate para realizar referencias cruzadas y validar información en libros diferentes. |
| [Contar filas en blanco en una hoja específica o en todas las hojas](count-blank-rows.md) | En este ejemplo se detecta si hay filas en blanco en hojas en las que se prevé que los datos estén presentes y, a continuación, se notifica el número de filas en blanco para su uso en un flujo de Power Automate. |
| [imágenes de gráficos y tablas de Email](email-images-chart-table.md) | En este ejemplo se usan scripts de Office y acciones de Power Automate para crear un gráfico y enviar ese gráfico como imagen por correo electrónico. |
| [Llamadas de captura externas](external-fetch-calls.md) | En este ejemplo se usa `fetch` para obtener información de GitHub para el script. |
| [Administración del modo de cálculo en Excel](excel-calculation.md) | En este ejemplo se muestra cómo usar el modo de cálculo y calcular métodos en Excel en la Web mediante scripts de Office. |
| [Mover filas entre tablas](move-rows-across-tables.md) | En este ejemplo se muestra cómo mover filas entre tablas guardando filtros y, a continuación, procesando y volviendo a aplicar los filtros. |
| [Salida de datos de Excel como JSON](get-table-data.md) | Esta solución muestra cómo generar datos de tabla de Excel como JSON para usarlos en Power Automate. |
| [Quitar hipervínculos de cada celda de una hoja de cálculo de Excel](remove-hyperlinks-from-cells.md) | En este ejemplo se borran todos los hipervínculos de la hoja de cálculo actual. |
| [Ejecutar un script en todos los archivos de Excel de una carpeta](automate-tasks-on-all-excel-files-in-folder.md) | Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa (también se puede usar para una carpeta de SharePoint). Realiza cálculos en los archivos de Excel, agrega formato e inserta un comentario que @mentions un compañero. |
| [Escribir un conjunto de datos grande](write-large-dataset.md) | En este ejemplo se muestra cómo enviar un intervalo grande como subranges más pequeños. |

## <a name="scenarios"></a>Escenarios

Los scripts de Office pueden automatizar partes de su rutina diaria. Estas tareas diarias a menudo existen en ecosistemas únicos, con libros de Excel configurados de maneras concretas. Estos ejemplos de escenarios más grandes demuestran estos casos de uso en el mundo real. Incluyen los scripts de Office y los libros, por lo que puede ver el escenario de un extremo a otro.

| Escenario | Detalles |
|---------|---------|
| [Analizar descargas web](../scenarios/analyze-web-downloads.md) | En este escenario se incluye un script que analiza los registros de tráfico web para determinar el país de origen de un usuario. Muestra las aptitudes del análisis de texto, el uso de subfunciones en scripts, la aplicación de formato condicional y el trabajo con tablas. |
| [Obtener y representar gráficamente datos del nivel de agua de NOAA](../scenarios/noaa-data-fetch.md) | En este escenario se usa un script de Office para extraer datos de un origen externo (la [base de datos Mareas y corrientes de NOAA) y](https://tidesandcurrents.noaa.gov/) representar gráficamente la información resultante. Resalta las aptitudes de uso `fetch` para obtener datos y usar gráficos. |
| [Calculadora de calificación](../scenarios/grade-calculator.md) | En este escenario se incluye un script que valida el registro de un instructor para las calificaciones de su clase. Muestra las aptitudes de comprobación de errores, formato de celda y expresiones regulares. |
| [Programar entrevistas en Teams](../scenarios/schedule-interviews-in-teams.md) | En este escenario se muestra cómo usar una hoja de cálculo de Excel para administrar las horas de reunión de las entrevistas y realizar un flujo a las reuniones de programación en Teams. |
| [Recordatorios de tareas](../scenarios/task-reminders.md) | En este escenario se usa un script de Office en un flujo de Power Automate para enviar recordatorios a compañeros de trabajo para actualizar el estado de un proyecto. Destaca las aptitudes de integración y transferencia de datos de Power Automate hacia y desde scripts. |

## <a name="community-contributions-and-fun-samples"></a>Contribuciones de la comunidad y ejemplos divertidos

¡Agradecemos [las contribuciones](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) de nuestra comunidad de Scripts de Office! No dude en crear una solicitud de incorporación de cambios para su revisión.

| Project | Detalles |
|---------|---------|
| [Juego de la vida](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | El blog "Ready Player Zero" de Yutao Huang en excel Tech Community incluye un script para modelar [*El juego de la vida de*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life) John Conway. |
| [Botón Reloj de fichar](../scenarios/punch-clock.md) | Este guión fue aportado por [Brian González](https://github.com/b-gonzalez). El escenario incluye un script y un botón de script que registra la hora actual. |
| [Animación de saludos de temporadas](community-seasons-greetings.md) | Este guión fue aportado por [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) en el espíritu de la temporada navideña. Es un divertido script que muestra un árbol de Navidad cantado en Excel en la Web con scripts de Office. |

## <a name="leave-a-comment"></a>Dejar un comentario

No dude en dejar un comentario, hacer una sugerencia o registrar un problema mediante la sección **Comentarios** en la parte inferior de la página de documentación del ejemplo específico.
