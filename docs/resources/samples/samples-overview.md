---
title: Ejemplos de scripts de Office
description: Ejemplos y escenarios de scripts de Office disponibles.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de0e99cbac7fcdeb1a3d3c43dd72ce53ed5847dd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571593"
---
# <a name="office-scripts-samples-and-scenarios"></a>Ejemplos y escenarios de scripts de Office

Esta sección contiene soluciones [de automatización basadas](../../overview/excel.md) en scripts de Office que ayudan a los usuarios finales a lograr la automatización de las tareas diarias. Contiene escenarios realistas a los que se enfrentan los usuarios empresariales y proporciona soluciones detalladas junto con vínculos de vídeo instructivo paso a paso.

Para cada uno de los proyectos de [Conceptos](#basics) básicos y Más allá de los [conceptos](#beyond-the-basics)básicos, consulte el código fuente, los vídeos paso a paso de [**YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)y mucho más.

En [Escenarios,](#scenarios)hemos incluido algunos ejemplos de escenarios más grandes que muestran casos de uso reales.

También agradecemos [las contribuciones de la comunidad.](#community-contributions)

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>Conceptos básicos

| Project | Detalles |
|---------|---------|
| [Conceptos básicos de scripting](../excel-samples.md) | Estos ejemplos muestran bloques de creación fundamentales para scripts de Office. |
| [Información básica sobre cómo usar el objeto Range en scripts de Office](range-basics.md) | En este artículo se muestran los conceptos básicos del uso del objeto Range y sus API. Este es un tema fundamental que se usará en todos los demás proyectos. |

## <a name="beyond-the-basics"></a>Más allá de los conceptos básicos

Consulte el siguiente proyecto completo que automatiza escenarios de ejemplo junto con scripts completos, archivos de Excel de ejemplo usados y [vídeos.](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Detalles |
|---------|---------|
| [Agregar comentarios en Excel](add-excel-comments.md) | En este ejemplo se muestra cómo agregar comentarios a una celda @mentioning un compañero. |
| [Contar filas en blanco en una hoja específica o en todas las hojas](count-blank-rows.md) | En este ejemplo se detecta si hay filas en blanco en hojas en las que se prevé que los datos estén presentes y, a continuación, se informa del recuento de filas en blanco para su uso en un flujo de Power Automate. |
| [Referencia cruzada y formato de un archivo de Excel](excel-cross-reference.md) | Esta solución muestra cómo se puede hacer referencia a dos archivos de Excel y dar formato con scripts de Office y Power Automate. |
| [Imágenes de tabla y gráfico de correo electrónico](email-images-chart-table.md) | En este ejemplo se usan scripts de Office y acciones de Power Automate para crear un gráfico y enviar ese gráfico como una imagen por correo electrónico. |
| [Filtrar tabla de Excel y obtener intervalo visible](filter-table-get-visible-range.md) | En este ejemplo se filtra una tabla de Excel y se devuelve el intervalo visible como un objeto JSON. Este JSON podría proporcionarse a un flujo de Power Automate como parte de una solución más grande. |
| [Generar un identificador único en un libro](document-number-generator.md) | Este escenario ayuda a un usuario a generar un número de documento único con un formato específico y agregar una entrada a un rango o tabla. |
| [Administrar el modo de cálculo en Excel](excel-calculation.md) | En este ejemplo se muestra cómo usar el modo de cálculo y calcular métodos en Excel en la web mediante scripts de Office. |
| [Combinar varias tablas de Excel en una sola tabla](copy-tables-combine.md) | En este ejemplo se combinan los datos de varias tablas de Excel en una sola tabla que incluye todas las filas. |
| [Mover filas entre tablas](move-rows-across-tables.md) | En este ejemplo se muestra cómo mover filas entre tablas guardando filtros y, a continuación, procesando y reaplicando los filtros. |
| [Salida de datos de Excel como JSON](get-table-data.md) | Esta solución muestra cómo generar datos de tabla de Excel como JSON para usarlos en Power Automate. |
| [Quitar hipervínculos de cada celda de una hoja de cálculo de Excel](remove-hyperlinks-from-cells.md) | En este ejemplo se borran todos los hipervínculos de la hoja de cálculo actual. |
| [Ejecutar un script en todos los archivos de Excel de una carpeta](automate-tasks-on-all-excel-files-in-folder.md) | Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa (también se puede usar para una carpeta de SharePoint). Realiza cálculos en los archivos de Excel, agrega formato e inserta un comentario que @mentions compañero. |
| [Enviar una reunión de Teams desde datos de Excel](send-teams-invite-from-excel-data.md) | Esta solución muestra cómo usar scripts de Office y acciones de Power Automate para seleccionar filas del archivo de Excel y usarlas para enviar una invitación a una reunión de Teams y, a continuación, actualizar Excel. |

## <a name="scenarios"></a>Escenarios

Los scripts de Office pueden automatizar partes de su rutina diaria. Estas tareas cotidianas a menudo existen en ecosistemas únicos, con libros de Excel que están configurados de maneras particulares. Estos ejemplos de escenarios más grandes muestran estos casos de uso reales. Incluyen los scripts de Office y los libros, por lo que puede ver el escenario de un extremo a otro.

| Escenario | Detalles |
|---------|---------|
| [Analizar descargas web](../scenarios/analyze-web-downloads.md) | En este escenario se incluye un script que analiza los registros de tráfico web para determinar el país de origen de un usuario. Muestra las habilidades del análisis de texto, el uso de subfunciones en scripts, la aplicación de formato condicional y el trabajo con tablas. |
| [Obtener y representar gráficamente datos del nivel de agua de NOAA](../scenarios/noaa-data-fetch.md) | Este escenario usa un script de Office para extraer datos de un origen externo (la base de datos de corrientes y [mareas de NOAA)](https://tidesandcurrents.noaa.gov/)y representar la información resultante. Destaca las habilidades de usar `fetch` para obtener datos y usar gráficos. |
| [Calculadora de calificación](../scenarios/grade-calculator.md) | En este escenario se incluye un script que valida el registro de un instructor para las calificaciones de su clase. Muestra las habilidades de comprobación de errores, formato de celda y expresiones regulares. |
| [Avisos de tareas](../scenarios/task-reminders.md) | Este escenario usa un script de Office en un flujo de Power Automate para enviar avisos a compañeros de trabajo para actualizar el estado de un proyecto. Destaca las habilidades de la integración de Power Automate y la transferencia de datos desde y hacia scripts. |

## <a name="community-contributions"></a>Contribuciones de la comunidad

Agradecemos [las contribuciones de](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) nuestra comunidad de scripts de Office. No dude en crear una solicitud de extracción para su revisión.

| Project | Detalles |
|---------|---------|
| [Animación de saludos de temporadas](community-seasons-greetings.md) | Este script fue contribuido por [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) en el ánimo de la temporada de vacaciones. Es un script divertido que muestra un árbol de Navidad cantando en Excel en la web con scripts de Office. |

## <a name="try-it-out"></a>Pruébelo

Estos ejemplos son de código abierto. Pruébalos tú mismo. Necesitarás una cuenta laboral o educativa de Microsoft desde el trabajo o la escuela con una licencia para la suscripción a Microsoft 365 (E3 o superior). Solo tienes que ir https://office.com a iniciar sesión en tu cuenta y empezar.

## <a name="leave-a-comment"></a>Dejar un comentario

No dude en dejar un comentario, hacer una sugerencia  o registrar un problema mediante la sección Comentarios en la parte inferior de la página de documentación del ejemplo específico.
