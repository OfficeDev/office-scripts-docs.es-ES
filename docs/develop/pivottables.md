---
title: Trabajar con tablas dinámicas en Scripts de Office
description: Obtenga información sobre el modelo de objetos para tablas dinámicas en la API de JavaScript de Scripts de Office.
ms.date: 04/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: a457c41bd1205f4e17636c43d7ba78addc80d0e4
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: es-ES
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572587"
---
# <a name="work-with-pivottables-in-office-scripts"></a>Trabajar con tablas dinámicas en Scripts de Office

Las tablas dinámicas permiten analizar rápidamente grandes colecciones de datos. Con su poder viene la complejidad. Las API de scripts de Office le permiten personalizar una tabla dinámica para satisfacer sus necesidades, pero el ámbito del conjunto de API hace que empezar sea un desafío. En este artículo se muestra cómo realizar tareas comunes de tabla dinámica y se explican clases y métodos importantes.

> [!NOTE]
> Para comprender mejor el contexto de los términos utilizados por las API, lea primero la documentación de la tabla dinámica de Excel. Comience con [Crear una tabla dinámica para analizar los datos de la hoja de cálculo](https://support.microsoft.com/office/a9a84538-bfe9-40a9-a8e9-f99134456576).

## <a name="object-model"></a>Modelo de objetos

:::image type="content" source="../images/pivottable-object-model.png" alt-text="Imagen simplificada de las clases, métodos y propiedades que se usan al trabajar con tablas dinámicas.":::

La [tabla dinámica](/javascript/api/office-scripts/excelscript/excelscript.pivottable) es el objeto central de las tablas dinámicas en la API de scripts de Office.

- El objeto [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) tiene una colección de todas las [tablas dinámicas](/javascript/api/office-scripts/excelscript/excelscript.pivottable). Cada [hoja de cálculo](/javascript/api/office-scripts/excelscript/excelscript.worksheet) también contiene una colección de tabla dinámica que es local a esa hoja.
- Una [tabla dinámica](/javascript/api/office-scripts/excelscript/excelscript.pivottable) contiene [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy). Una jerarquía se puede considerar como una columna de una tabla.
- [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) se puede agregar como filas o columnas ([RowColumnPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.rowcolumnpivothierarchy)), datos ([DataPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.datapivothierarchy)) o filtros ([FilterPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)).
- Cada [pivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) contiene exactamente un [campo dinámico](/javascript/api/office-scripts/excelscript/excelscript.pivotfield). Las estructuras de tabla dinámica fuera de Excel pueden contener varios campos por jerarquía, por lo que este diseño existe para admitir opciones futuras. En el caso de los scripts de Office, los campos y las jerarquías se asignan a la misma información.
- Un [campo dinámico](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) contiene varios [elementos PivotItem](/javascript/api/office-scripts/excelscript/excelscript.pivotitem). Cada objeto PivotItem es un valor único en el campo. Piense en cada elemento como un valor de la columna de tabla. Los elementos también pueden ser valores agregados, como sumas, si el campo se usa para los datos.
- [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) define cómo se muestran los [pivotfields](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) y [pivotitems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem).
- [PivotFilters filtra los](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters) datos de la [tabla dinámica](/javascript/api/office-scripts/excelscript/excelscript.pivottable) con criterios diferentes.

Examine cómo funcionan estas relaciones en la práctica. En los datos siguientes se describen las ventas de frutas de varias granjas de servidores. Es la base de todos los ejemplos de este artículo. Use [pivottable-sample.xlsx](pivottable-sample.xlsx) para seguir.

:::image type="content" source="../images/pivottable-raw-data.png" alt-text="Colección de ventas de frutas de diferentes tipos de granjas.":::

## <a name="create-a-pivottable-with-fields"></a>Creación de una tabla dinámica con campos

Las tablas dinámicas se crean con referencias a datos existentes. Tanto los intervalos como las tablas pueden ser el origen de una tabla dinámica. También necesitan un lugar para existir en el libro. Dado que el tamaño de una tabla dinámica es dinámico, solo se especifica la esquina superior izquierda del intervalo de destino.

El siguiente fragmento de código crea una tabla dinámica basada en un intervalo de datos. La tabla dinámica no tiene jerarquías, por lo que los datos aún no se agrupan de ninguna manera.

```typescript
  const dataSheet = workbook.getWorksheet("Data");
  const pivotSheet = workbook.getWorksheet("Pivot");

  const farmPivot = pivotSheet.addPivotTable(
    "Farm Pivot", /* The name of the PivotTable. */
    dataSheet.getUsedRange(), /* The source data range. */
    pivotSheet.getRange("A1") /* The location to put the new PivotTable. */);
```

:::image type="content" source="../images/pivottable-empty.png" alt-text="Una tabla dinámica denominada &quot;Farm Pivot&quot; sin jerarquías.":::

### <a name="hierarchies-and-fields"></a>Jerarquías y campos

Las tablas dinámicas se organizan mediante jerarquías. Esas jerarquías se usan para dinamizar los datos cuando se agregan como un tipo específico de jerarquía. Hay cuatro tipos de jerarquías.

- **Fila**: muestra elementos en filas horizontales.
- **Columna**: muestra elementos en columnas verticales.
- **Datos**: muestra agregados de valores basados en las filas y columnas.
- **Filtro**: agrega o quita elementos de la tabla dinámica.

Una tabla dinámica puede tener tantos o como pocos de sus campos asignados a estas jerarquías específicas. Una tabla dinámica necesita al menos una jerarquía de datos para mostrar datos numéricos resumidos y al menos una fila o columna en la que pivotar ese resumen. El siguiente fragmento de código agrega dos jerarquías de filas y dos jerarquías de datos.

```typescript
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Farm"));
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Type"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold at Farm"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold Wholesale"));
```

:::image type="content" source="../images/pivottable-data-hierarchy.png" alt-text="Tabla dinámica que muestra las ventas totales de diferentes frutas basadas en la granja de la que proceden.":::

## <a name="layout-ranges"></a>Intervalos de diseño

Cada parte de la tabla dinámica se asigna a un intervalo. Esto permite que el script obtenga datos de la tabla dinámica para usarlos más adelante en el script o para devolverlos en un [flujo de Power Automate](power-automate-integration.md). Se obtiene acceso a estos intervalos a través del objeto [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) adquirido de `PivotTable.getLayout()`. En el diagrama siguiente se muestran los intervalos devueltos por los métodos de `PivotLayout`.

:::image type="content" source="../images/pivottable-layout-breakdown.png" alt-text="Diagrama en el que se muestran las secciones de una tabla dinámica devueltas por las funciones get range del diseño.":::

## <a name="filters-and-slicers"></a>Filtros y segmentaciones de datos

Hay tres maneras de filtrar una tabla dinámica.

- [FilterPivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters)
- [Slicers](/javascript/api/office-scripts/excelscript/excelscript.slicer)

### <a name="filterpivothierarchies"></a>FilterPivotHierarchies

`FilterPivotHierarchies` agregue una jerarquía adicional para filtrar cada fila de datos. Cualquier fila con un elemento filtrado se excluye de la tabla dinámica y sus resúmenes. Dado que estos filtros se basan en elementos, solo funcionan en valores discretos. Si "Clasificación" es una jerarquía de filtros en nuestro ejemplo, los usuarios pueden seleccionar los valores "Orgánico" y "Convencional" para el filtro. De forma similar, si se selecciona "Cajas vendidas al por mayor", las opciones de filtro serían los números individuales, como 120 y 150, en lugar de rangos numéricos.

`FilterPivotHierarchies` se crean con todos los valores seleccionados. Esto significa que no se filtra nada hasta que el usuario interactúa manualmente con el control de filtro o se establece en `PivotManualFilter` el campo que pertenece a `FilterPivotHierarchy`.

El siguiente fragmento de código agrega "Clasificación" como jerarquía de filtros.

```typescript
  farmPivot.addFilterHierarchy(farmPivot.getHierarchy("Classification"));
```

:::image type="content" source="../images/pivottable-filter-hierarchy.png" alt-text="Control de filtro que usa &quot;Classification&quot; para una tabla dinámica.":::

### <a name="pivotfilters"></a>PivotFilters

El `PivotFilters` objeto es una colección de filtros aplicados a un único campo. Dado que cada jerarquía tiene exactamente un campo, siempre debe usar el primer campo en `PivotHierarchy.getFields()` al aplicar filtros. Hay cuatro tipos de filtro.

- **Filtro de fecha**: filtrado basado en fechas del calendario.
- **Filtro de etiqueta**: filtrado de comparación de texto.
- **Filtro manual**: filtrado de entrada personalizado.
- **Filtro de valor**: filtrado de comparación de números. Esto compara los elementos de la jerarquía asociada con los valores de una jerarquía de datos especificada.

Normalmente, solo se crea uno de los cuatro tipos de filtros y se aplica al campo. Si el script intenta usar filtros incompatibles, se produce un error con el texto "El argumento no es válido o falta o tiene un formato incorrecto".

El siguiente fragmento de código agrega dos filtros. El primero es un filtro manual que selecciona los elementos de una jerarquía de filtros "Clasificación" existente. El segundo filtro quita las granjas que tienen menos de 300 "Cajas vendidas al por mayor". Tenga en cuenta que esto filtra la "suma" de esas granjas de servidores, no las filas individuales de los datos originales.

```typescript
  const classificationField = farmPivot.getFilterHierarchy("Classification").getFields()[0];
  classificationField.applyFilter({
    manualFilter: { 
      selectedItems: ["Organic"] /* The included items. */
    }
  });

  const farmField = farmPivot.getHierarchy("Farm").getFields()[0];
  farmField.applyFilter({
    valueFilter: {
      condition: ExcelScript.ValueFilterCondition.greaterThan, /* The relationship of the value to the comparator. */
      comparator: 300, /* The value to which items are compared. */
      value: "Sum of Crates Sold Wholesale" /* The name of the data hierarchy. Note the "Sum of" prefix. */
      }
  });
```

:::image type="content" source="../images/pivottable-filters.png" alt-text="Una tabla dinámica después de aplicar el filtro de valor y el filtro manual.":::

### <a name="slicers"></a>Segmentación de datos

[Las segmentaciones filtran](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) los datos de una tabla dinámica (o una tabla estándar). Son objetos que se pueden mover en la hoja de cálculo que permiten filtrar rápidamente las selecciones. Una segmentación funciona de forma similar al filtro manual y `PivotFilterHierarchy`. Los elementos de `PivotField` se alternan para incluirlos o excluirlos de la tabla dinámica.

El siguiente fragmento de código agrega una segmentación para el campo "Type". Establece los elementos seleccionados para que sean "Lemon" y "Lime" y, a continuación, mueve la segmentación de 400 píxeles a la izquierda.

```typescript
  const fruitSlicer = pivotSheet.addSlicer(
    farmPivot, /* The table or PivotTale to be sliced. */
    farmPivot.getHierarchy("Type").getFields()[0] /* What source to use as the slicer options. */
  );
  fruitSlicer.selectItems(["Lemon", "Lime"]);
  fruitSlicer.setLeft(400);
```

:::image type="content" source="../images/slicer.png" alt-text="Una segmentación de datos que filtra datos en una tabla dinámica.":::

## <a name="see-also"></a>Vea también

- [Conceptos básicos de los Scripts de Office en Excel en la web](scripting-fundamentals.md)
- [Referencia de API de scripts de Office](/javascript/api/office-scripts/overview)
