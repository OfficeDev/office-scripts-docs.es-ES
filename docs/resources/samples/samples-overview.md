---
title: ejemplos de scripts de Office
description: Escenarios y ejemplos de scripts de Office disponibles.
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c9bbe9b6f7eb8abad2995dac72ccf636d585d69
ms.sourcegitcommit: e6428a5214fa38aef036a952a0e3c09dbf6e4d3e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/28/2022
ms.locfileid: "65109164"
---
# <a name="office-scripts-samples-and-scenarios"></a>escenarios y ejemplos de scripts de Office

Esta sección contiene Office soluciones de automatización [basadas en scripts](../../overview/excel.md) que ayudan a los usuarios finales a lograr la automatización de las tareas diarias. Contiene escenarios realistas a los que se enfrentan los usuarios empresariales y proporciona soluciones detalladas junto con vínculos de vídeo de instrucciones paso a paso.

Para cada uno de los proyectos de [Conceptos básicos](#basics) y [Más allá de los conceptos básicos](#beyond-the-basics), consulte el código fuente, [**los vídeos de YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0) paso a paso y mucho más.

En Escenarios, hemos incluido algunos [ejemplos](#scenarios) de escenarios más grandes que muestran casos de uso reales.

También agradecemos [las contribuciones de la comunidad](#community-contributions-and-fun-samples).

## <a name="basics"></a>Conceptos básicos

| Project | Detalles |
|---------|---------|
| [Conceptos básicos de scripting](../excel-samples.md) | En estos ejemplos se muestran los bloques de creación fundamentales para scripts de Office. |
| [Agregar comentarios en Excel](add-excel-comments.md) | En este ejemplo se agregan comentarios a una celda, incluido @mentioning un compañero. |
| [Agregar imágenes a un libro](add-image-to-workbook.md) | En este ejemplo se agrega una imagen a un libro y se copia una imagen entre hojas.|
| [Copia de varias tablas Excel en una sola tabla](copy-tables-combine.md) | Este ejemplo combina datos de varias tablas de Excel en una sola tabla que incluye todas las filas. |
| [Crear una tabla de contenido de libro](table-of-contents.md) | En este ejemplo se crea una tabla de contenido con vínculos a cada hoja de cálculo. |

## <a name="beyond-the-basics"></a>Más allá de los aspectos básicos

Consulte el siguiente proyecto de un extremo a otro que automatiza escenarios de ejemplo junto con scripts completos, archivos de Excel de ejemplo usados y [vídeos (hospedados en YouTube).](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Detalles |
|---------|---------|
| [Combinar hojas de cálculo en un solo libro](combine-worksheets-into-single-workbook.md) | En este ejemplo se usan scripts de Office y Power Automate para extraer datos de otros libros en un solo libro. |
| [Conversión de archivos CSV en libros de Excel](convert-csv.md) | En este ejemplo se usan scripts de Office y Power Automate para crear archivos .xlsx a partir de archivos .csv. |
| [Libros de referencia cruzada](excel-cross-reference.md) | En este ejemplo se usan scripts de Office y Power Automate para realizar referencias cruzadas y validar información en libros diferentes. |
| [Contar filas en blanco en una hoja específica o en todas las hojas](count-blank-rows.md) | En este ejemplo se detecta si hay filas en blanco en hojas en las que se prevé que los datos estén presentes y, a continuación, se notifica el número de filas en blanco para su uso en un flujo de Power Automate. |
| [Gráfico de correo electrónico e imágenes de tabla](email-images-chart-table.md) | En este ejemplo se usan scripts de Office y acciones de Power Automate para crear un gráfico y enviar ese gráfico como imagen por correo electrónico. |
| [Llamadas de captura externas](external-fetch-calls.md) | En este ejemplo se usa `fetch` para obtener información de GitHub para el script. |
| [Filtrar Excel tabla y obtener el intervalo visible](filter-table-get-visible-range.md) | Este ejemplo filtra una tabla Excel y devuelve el intervalo visible como un objeto JSON. Este JSON se podría proporcionar a un flujo de Power Automate como parte de una solución más grande. |
| [Administración del modo de cálculo en Excel](excel-calculation.md) | En este ejemplo se muestra cómo usar el modo de cálculo y calcular métodos en Excel en la Web mediante scripts de Office. |
| [Mover filas entre tablas](move-rows-across-tables.md) | En este ejemplo se muestra cómo mover filas entre tablas guardando filtros y, a continuación, procesando y volviendo a aplicar los filtros. |
| [Salida Excel datos como JSON](get-table-data.md) | Esta solución muestra cómo generar Excel datos de tabla como JSON para usarlos en Power Automate. |
| [Quitar hipervínculos de cada celda de una hoja de cálculo de Excel](remove-hyperlinks-from-cells.md) | En este ejemplo se borran todos los hipervínculos de la hoja de cálculo actual. |
| [Ejecutar un script en todos los archivos de Excel de una carpeta](automate-tasks-on-all-excel-files-in-folder.md) | Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa (también se puede usar para una carpeta SharePoint). Realiza cálculos en los archivos Excel, agrega formato e inserta un comentario que @mentions un compañero. |
| [Escribir un conjunto de datos grande](write-large-dataset.md) | En este ejemplo se muestra cómo enviar un intervalo grande como subranges más pequeños. |

## <a name="scenarios"></a>Escenarios

Office Scripts puede automatizar partes de la rutina diaria. Estas tareas diarias a menudo existen en ecosistemas únicos, con Excel libros configurados de maneras particulares. Estos ejemplos de escenarios más grandes demuestran estos casos de uso en el mundo real. Incluyen tanto los scripts de Office como los libros, para que pueda ver el escenario de un extremo a otro.

| Escenario | Detalles |
|---------|---------|
| [Analizar descargas web](../scenarios/analyze-web-downloads.md) | En este escenario se incluye un script que analiza los registros de tráfico web para determinar el país de origen de un usuario. Muestra las aptitudes del análisis de texto, el uso de subfunciones en scripts, la aplicación de formato condicional y el trabajo con tablas. |
| [Obtener y representar gráficamente datos del nivel de agua de NOAA](../scenarios/noaa-data-fetch.md) | En este escenario se usa un script de Office para extraer datos de un origen externo (la [base de datos Mareas y corrientes de NOAA) y](https://tidesandcurrents.noaa.gov/) representar gráficamente la información resultante. Resalta las aptitudes de uso `fetch` para obtener datos y usar gráficos. |
| [Calculadora de calificación](../scenarios/grade-calculator.md) | En este escenario se incluye un script que valida el registro de un instructor para las calificaciones de su clase. Muestra las aptitudes de comprobación de errores, formato de celda y expresiones regulares. |
| [Programar entrevistas en Teams](../scenarios/schedule-interviews-in-teams.md) | En este escenario se muestra cómo usar una hoja de cálculo de Excel para administrar las horas de reunión de las entrevistas y realizar un flujo a las reuniones programadas en Teams. |
| [Recordatorios de tareas](../scenarios/task-reminders.md) | En este escenario se usa un script de Office en un flujo de Power Automate para enviar recordatorios a compañeros de trabajo para actualizar el estado de un proyecto. Resalta las aptitudes de Power Automate integración y transferencia de datos hacia y desde scripts. |

## <a name="community-contributions-and-fun-samples"></a>Community contribuciones y ejemplos divertidos

¡Agradecemos [las contribuciones](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) de nuestra comunidad Office Scripts! No dude en crear una solicitud de incorporación de cambios para su revisión.

| Project | Detalles |
|---------|---------|
| [Juego de la vida](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | El blog "Ready Player Zero" de Yutao Huang en el Excel Tech Community incluye un script para modelar [*The Game of Life de*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life) John Conway. |
| [Botón de reloj de punzonado](../scenarios/punch-clock.md) | Este guión fue aportado por [Brian González](https://github.com/b-gonzalez). El escenario incluye un script y un botón de script que registra la hora actual. |
| [Animación de saludos de temporadas](community-seasons-greetings.md) | Este guión fue aportado por [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) en el espíritu de la temporada navideña. Es un divertido script que muestra un árbol de Navidad cantando en Excel en la Web con scripts Office. |

## <a name="try-it-out"></a>Pruébelo

Estos ejemplos se código abierto. Pruébelo usted mismo. Necesitará una cuenta profesional o educativa de Microsoft profesional o educativa con una licencia para Microsoft 365 suscripción (E3 o superior). Vaya a https://office.com para iniciar sesión en su cuenta y empezar a trabajar.

## <a name="leave-a-comment"></a>Dejar un comentario

No dude en dejar un comentario, hacer una sugerencia o registrar un problema mediante la sección **Comentarios** en la parte inferior de la página de documentación del ejemplo específico.
