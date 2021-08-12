---
title: Office Ejemplos de scripts
description: Disponible Office ejemplos de scripts y escenarios.
ms.date: 07/19/2021
localization_priority: Normal
ms.openlocfilehash: 508c9fa61d6816faf3fbe740bdf1d587b127b064e9918aac2d789ff2d4faccff
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846352"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Ejemplos y escenarios de scripts

En esta sección se [Office soluciones](../../overview/excel.md) de automatización basadas en scripts que ayudan a los usuarios finales a lograr la automatización de las tareas diarias. Contiene escenarios realistas a los que se enfrentan los usuarios empresariales y proporciona soluciones detalladas junto con vínculos de vídeo instructivo paso a paso.

Para cada uno de los proyectos de [Conceptos](#basics) básicos y Más allá de los [conceptos](#beyond-the-basics)básicos, consulte el código fuente, los vídeos paso a paso de [**YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)y mucho más.

En [Escenarios,](#scenarios)hemos incluido algunos ejemplos de escenarios más grandes que muestran casos de uso reales.

También agradecemos [las contribuciones de la comunidad.](#community-contributions-and-fun-samples)

## <a name="basics"></a>Datos básicos

| Project | Detalles |
|---------|---------|
| [Conceptos básicos de scripting](../excel-samples.md) | Estos ejemplos muestran bloques de creación fundamentales para Office scripts. |
| [Agregar comentarios en Excel](add-excel-comments.md) | En este ejemplo se agregan comentarios a una celda @mentioning un compañero. |
| [Agregar imágenes a un libro](add-image-to-workbook.md) | En este ejemplo se agrega una imagen a un libro y se copia una imagen entre hojas.|
| [Copiar varias Excel en una sola tabla](copy-tables-combine.md) | En este ejemplo se combinan los datos de varias Excel tablas en una sola tabla que incluye todas las filas. |

## <a name="beyond-the-basics"></a>Más allá de los aspectos básicos

Consulte el siguiente proyecto completo que automatiza escenarios de ejemplo junto con scripts completos, archivos Excel de ejemplo usados y vídeos (hospedados en [YouTube).](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Detalles |
|---------|---------|
| [Convertir archivos CSV en Excel libros](convert-csv.md) | En este ejemplo se Office scripts y Power Automate para crear .xlsx archivos .csv archivos. |
| [Libros de referencia cruzada](excel-cross-reference.md) | En este ejemplo se Office scripts y Power Automate para hacer referencia cruzada y validar información en diferentes libros. |
| [Contar filas en blanco en una hoja específica o en todas las hojas](count-blank-rows.md) | En este ejemplo se detecta si hay filas en blanco en hojas en las que se prevé que los datos estén presentes y, a continuación, se informa del recuento de filas en blanco para su uso en un flujo Power Automate datos. |
| [Imágenes de tabla y gráfico de correo electrónico](email-images-chart-table.md) | En este ejemplo se Office scripts y Power Automate acciones para crear un gráfico y enviar dicho gráfico como una imagen por correo electrónico. |
| [Llamadas de captura externa](external-fetch-calls.md) | En este ejemplo `fetch` se usa para obtener información GitHub para el script. |
| [Filtrar Excel tabla y obtener rango visible](filter-table-get-visible-range.md) | En este ejemplo se filtra Excel tabla y se devuelve el intervalo visible como un objeto JSON. Este JSON podría proporcionarse a un flujo Power Automate como parte de una solución más grande. |
| [Administrar el modo de cálculo en Excel](excel-calculation.md) | En este ejemplo se muestra cómo usar el modo de cálculo y calcular métodos en Excel en la Web usar Office scripts. |
| [Mover filas entre tablas](move-rows-across-tables.md) | En este ejemplo se muestra cómo mover filas entre tablas guardando filtros y, a continuación, procesando y reaplicando los filtros. |
| [Datos Excel salida como JSON](get-table-data.md) | Esta solución muestra cómo generar datos Excel tabla como JSON para usarlos en Power Automate. |
| [Quitar hipervínculos de cada celda de una hoja Excel hoja de cálculo](remove-hyperlinks-from-cells.md) | En este ejemplo se borran todos los hipervínculos de la hoja de cálculo actual. |
| [Ejecutar un script en todos los archivos de Excel de una carpeta](automate-tasks-on-all-excel-files-in-folder.md) | Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa (también se puede usar para una SharePoint carpeta). Realiza cálculos en los archivos Excel, agrega formato e inserta un comentario que @mentions compañero. |
| [Escribir un conjunto de datos grande](write-large-dataset.md) | En este ejemplo se muestra cómo enviar un rango grande como subranges más pequeños. |

## <a name="scenarios"></a>Escenarios

Office Los scripts pueden automatizar partes de la rutina diaria. Estas tareas cotidianas a menudo existen en ecosistemas únicos, con Excel libros que se establecen de maneras particulares. Estos ejemplos de escenarios más grandes muestran estos casos de uso reales. Incluyen los scripts de Office y los libros, para que pueda ver el escenario de un extremo a otro.

| Escenario | Detalles |
|---------|---------|
| [Analizar descargas web](../scenarios/analyze-web-downloads.md) | En este escenario se incluye un script que analiza los registros de tráfico web para determinar el país de origen de un usuario. Muestra las habilidades del análisis de texto, el uso de subfunciones en scripts, la aplicación de formato condicional y el trabajo con tablas. |
| [Obtener y representar gráficamente datos del nivel de agua de NOAA](../scenarios/noaa-data-fetch.md) | Este escenario usa un script Office para extraer datos de un origen externo (la base de datos de corrientes y [mareas de NOAA)](https://tidesandcurrents.noaa.gov/)y representar la información resultante. Destaca las habilidades de usar `fetch` para obtener datos y usar gráficos. |
| [Calculadora de calificación](../scenarios/grade-calculator.md) | En este escenario se incluye un script que valida el registro de un instructor para las calificaciones de su clase. Muestra las habilidades de comprobación de errores, formato de celda y expresiones regulares. |
| [Programar entrevistas en Teams](../scenarios/schedule-interviews-in-teams.md) | En este escenario se muestra cómo usar una hoja de Excel para administrar las horas de reunión de las entrevistas y hacer un flujo a las reuniones programadas en Teams. |
| [Avisos de tareas](../scenarios/task-reminders.md) | Este escenario usa un script Office en un flujo de Power Automate para enviar avisos a compañeros de trabajo para actualizar el estado de un proyecto. Destaca las habilidades de Power Automate integración y transferencia de datos desde y hacia scripts. |

## <a name="community-contributions-and-fun-samples"></a>Community contribuciones y muestras divertidas

¡Agradecemos [las contribuciones](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) de nuestra Office scripts! No dude en crear una solicitud de extracción para su revisión.

| Project | Detalles |
|---------|---------|
| [Juego de vida](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | El blog "Ready Player Zero" de Yutao Huang en Excel Tech Community incluye un script para modelar [*El juego*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life)de la vida de John Conway. |
| [Animación de saludos de temporadas](community-seasons-greetings.md) | Este script fue contribuido por [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) en el ánimo de la temporada de vacaciones. Es un script divertido que muestra un árbol de Navidad cantando en Excel en la Web usando Office scripts. |

## <a name="try-it-out"></a>Pruébelo

Estos ejemplos son de código abierto. Pruébalos tú mismo. Necesitarás una cuenta laboral o educativa de Microsoft desde el trabajo o la escuela con una licencia para Microsoft 365 suscripción (E3 o superior). Solo tienes que ir https://office.com a iniciar sesión en tu cuenta y empezar.

## <a name="leave-a-comment"></a>Dejar un comentario

No dude en dejar un comentario, hacer una sugerencia  o registrar un problema mediante la sección Comentarios en la parte inferior de la página de documentación del ejemplo específico.
