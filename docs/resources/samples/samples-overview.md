---
title: Office Ejemplos de scripts
description: Disponible Office scripts ejemplos y escenarios.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0ea9a8a8986681fca0e45784e2923c1d3b34576d
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545712"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Scripts muestras y escenarios

Esta sección contiene soluciones de automatización basadas en [scripts Office](../../overview/excel.md) que ayudan a los usuarios finales a lograr la automatización de las tareas diarias. Contiene escenarios realistas a los que se enfrentan los usuarios empresariales y proporciona soluciones detalladas junto con enlaces de vídeo instructivo paso a paso.

Para cada uno de los proyectos en [Conceptos básicos](#basics) y más allá de [lo básico,](#beyond-the-basics)echa un vistazo al código fuente, vídeos de [**YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)paso a paso y mucho más.

En [Escenarios,](#scenarios)hemos incluido algunos ejemplos de escenarios más grandes que demuestran casos de uso del mundo real.

También damos la bienvenida [a las contribuciones de la comunidad.](#community-contributions-and-fun-samples)

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>Conceptos básicos

| Project | Detalles |
|---------|---------|
| [Conceptos básicos de scripting](../excel-samples.md) | Estos ejemplos muestran bloques de creación fundamentales para scripts de Office. |
| [Añadir comentarios en Excel](add-excel-comments.md) | En este ejemplo se agregan comentarios a una celda, incluida @mentioning un colega. |
| [Agregar imágenes a un libro](add-image-to-workbook.md) | En este ejemplo se agrega una imagen a un libro de trabajo y se copia una imagen en todas las hojas.|
| [Copie varias tablas de Excel en una sola tabla](copy-tables-combine.md) | Este ejemplo combina datos de varias tablas de Excel en una sola tabla que incluye todas las filas. |

## <a name="beyond-the-basics"></a>Más allá de los aspectos básicos

Consulte el siguiente proyecto de extremo a extremo que automatiza escenarios de ejemplo junto con scripts completos, archivos de Excel de ejemplo utilizados y [vídeos (alojados en YouTube).](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Detalles |
|---------|---------|
| [Cuente filas en blanco en una hoja específica o en todas las hojas](count-blank-rows.md) | Este ejemplo detecta si hay filas en blanco en hojas donde se prevén que los datos estén presentes y, a continuación, se informa del recuento de filas en blanco para su uso en un flujo de Power Automate. |
| [Gráfico de correo electrónico e imágenes de tabla](email-images-chart-table.md) | En este ejemplo se usan Office scripts y acciones de Power Automate para crear un gráfico y enviar ese gráfico como una imagen por correo electrónico. |
| [Llamadas de captura externas](external-fetch-calls.md) | Este ejemplo se utiliza `fetch` para obtener información de GitHub para el script. |
| [Filtre Excel tabla y obtenga un rango visible](filter-table-get-visible-range.md) | En este ejemplo se filtra una tabla Excel y se devuelve el intervalo visible como un objeto JSON. Este JSON podría proporcionarse a un flujo Power Automate como parte de una solución más grande. |
| [Administrar el modo de cálculo en Excel](excel-calculation.md) | En este ejemplo se muestra cómo utilizar el modo de cálculo y calcular métodos en Excel en la Web mediante scripts de Office. |
| [Mover filas a través de tablas](move-rows-across-tables.md) | En este ejemplo se muestra cómo mover filas entre tablas guardando filtros y, a continuación, procesando y volver a aplicar los filtros. |
| [Salida Excel datos como JSON](get-table-data.md) | Esta solución muestra cómo generar Excel datos de tabla como JSON para usar en Power Automate. |
| [Elimine los hipervínculos de cada celda de una hoja de trabajo Excel](remove-hyperlinks-from-cells.md) | Este ejemplo borra todos los hipervínculos de la hoja de cálculo actual. |
| [Ejecutar un script en todos los archivos de Excel de una carpeta](automate-tasks-on-all-excel-files-in-folder.md) | Este proyecto realiza un conjunto de tareas de automatización en todos los archivos situados en una carpeta de OneDrive para la Empresa (también se puede utilizar para una carpeta SharePoint). Realiza cálculos en los archivos Excel, agrega formato e inserta un comentario que @mentions un colega. |
| [Escribir un conjunto de datos grande](write-large-dataset.md) | En este ejemplo se muestra cómo enviar un rango grande como subranges más pequeños. |

## <a name="scenarios"></a>Escenarios

Office Los scripts pueden automatizar partes de su rutina diaria. Estas tareas diarias a menudo existen en ecosistemas únicos, con Excel libros de trabajo que se configuran de maneras particulares. Estas muestras de escenarios más grandes demuestran estos casos de uso en el mundo real. Incluyen tanto los scripts de Office como los libros de trabajo, por lo que puede ver el escenario de extremo a extremo.

| Escenario | Detalles |
|---------|---------|
| [Analizar descargas web](../scenarios/analyze-web-downloads.md) | Este escenario cuenta con un script que analiza los registros de tráfico web para determinar el país de origen de un usuario. Muestra las habilidades del análisis de texto, el uso de subfunciones en scripts, la aplicación de formato condicional y el trabajo con tablas. |
| [Obtener y representar gráficamente datos del nivel de agua de NOAA](../scenarios/noaa-data-fetch.md) | Este escenario utiliza un script Office para extraer datos de un origen externo (la base de datos mareas y corrientes de la [NOAA)](https://tidesandcurrents.noaa.gov/)y graficar la información resultante. Destaca las habilidades de usar `fetch` para obtener datos y usar gráficos. |
| [Calculadora de calificación](../scenarios/grade-calculator.md) | Este escenario cuenta con un script que valida el registro de un instructor para las calificaciones de su clase. Muestra las habilidades de comprobación de errores, formato de celda y expresiones regulares. |
| [Recordatorios de tareas](../scenarios/task-reminders.md) | Este escenario usa un script Office en un flujo de Power Automate para enviar recordatorios a los compañeros de trabajo para actualizar el estado de un proyecto. Destaca las habilidades de Power Automate integración y transferencia de datos hacia y desde scripts. |

## <a name="community-contributions-and-fun-samples"></a>Community contribuciones y muestras divertidas

¡Damos la bienvenida [a las contribuciones](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) de nuestra comunidad de scripts Office! Siéntase libre de crear una solicitud de extracción para su revisión.

| Project | Detalles |
|---------|---------|
| [Juego de la vida](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | El blog "Ready Player Zero" de Yutao Huang en el Excel Tech Community incluye un guión para modelar [*The Game of Life de*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life)John Conway. |
| [Temporadas saludando animación](community-seasons-greetings.md) | Este guión fue contribuido por [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) en el espíritu de la temporada de vacaciones! Es un guión divertido que muestra un árbol de Navidad cantando en Excel en la Web usando Office Scripts. |

## <a name="try-it-out"></a>Pruébelo

Estas muestras son de código abierto. Pruébalos tú mismo. Necesitarás una cuenta profesional o educativa de Microsoft desde el trabajo o la escuela con una licencia para Microsoft 365 suscripción (E3 o superior). Sólo tienes que https://office.com dirigirte para iniciar sesión en tu cuenta y empezar.

## <a name="leave-a-comment"></a>Deja un comentario

Siéntase libre de dejar un comentario, hacer una sugerencia o registrar un problema mediante la sección **Comentarios** en la parte inferior de la página de documentación del ejemplo específico.
