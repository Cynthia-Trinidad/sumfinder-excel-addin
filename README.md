# SumFinder — Excel Add-in

Encuentra qué combinaciones de celdas suman un monto específico.
Recibe 3 opciones: la más cercana al objetivo, una con menos y una con más.

---

## ¿Qué hace?

1. Seleccionas un rango de celdas en Excel
2. Escribes el monto que quieres alcanzar (o lo tomas de una celda)
3. SumFinder encuentra combinaciones y te da 3 opciones:
   - 🎯 **Más cercano** — la suma más exacta posible
   - ▼ **Un poco menos** — combinación que queda por debajo
   - ▲ **Un poco más** — combinación que queda por encima
4. Para cada resultado puedes: **Seleccionar**, **Copiar** o **Cortar** las celdas

---

## Estructura del proyecto

```
excel-addin/
├── manifest.xml          ← Le dice a Excel qué es este add-in
├── package.json          ← Dependencias y scripts
└── src/
    ├── taskpane.html     ← Interfaz del panel lateral
    ├── taskpane.js       ← Lógica principal (algoritmo + Office.js)
    └── commands.html     ← Requerido por el manifest
```

---

## Cómo instalar y ejecutar (desarrollo local)

### Requisitos
- Node.js 16+
- Excel Desktop (Windows o Mac) o Excel Online

### Pasos

**1. Instalar dependencias**
```bash
npm install
```

**2. Instalar certificados SSL locales** (necesario para que Excel confíe en localhost)
```bash
npx office-addin-dev-certs install --machine
```

**3. Levantar el servidor local**
```bash
npm start
```
Esto sirve los archivos en `https://localhost:3000`

**4. Cargar el add-in en Excel**

*En Excel Desktop (Windows):*
- Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza
- Catálogos de complementos de confianza → agrega la carpeta del proyecto
- O usa: Insertar → Obtener complementos → Mis complementos → Cargar desde archivo → selecciona `manifest.xml`

*En Excel Online:*
- Insertar → Obtener complementos → Mis complementos → Cargar complemento → sube `manifest.xml`

---

## Cómo agregar nuevas funciones en el futuro

El proyecto está diseñado para escalar. Cada función nueva se agrega en `taskpane.js`.

**Ejemplos de funciones que podrías agregar:**

### Función: Resaltar celdas con color
```javascript
async function highlightCells(cells, color = "#FFE066") {
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.activeWorksheet;
    const range = ws.getRanges(cells.join(","));
    range.format.fill.color = color;
    await ctx.sync();
  });
}
```

### Función: Exportar resultado a una nueva hoja
```javascript
async function exportResultToSheet(result) {
  await Excel.run(async (ctx) => {
    const newSheet = ctx.workbook.worksheets.add("SumFinder_Resultado");
    newSheet.getRange("A1").values = [["Celda", "Valor"]];
    result.cells.forEach((addr, i) => {
      newSheet.getRange(`A${i+2}`).values = [[addr, result.sum]];
    });
    newSheet.activate();
    await ctx.sync();
  });
}
```

Para agregar un botón en la UI, simplemente añades en `taskpane.html`:
```html
<button class="btn-action" onclick="highlightCells(result.cells)">🎨 Resaltar</button>
```

---

## Algoritmo

Para rangos grandes (200-500 celdas) es imposible probar todas las combinaciones
posibles (2^500 ≈ infinito). SumFinder usa una **heurística greedy + perturbaciones**:

1. Ordena los valores por magnitud y aplica una estrategia greedy
2. Prueba múltiples órdenes aleatorios con semillas distintas
3. Aplica perturbaciones (swap/add/remove) sobre las mejores soluciones
4. Filtra los resultados en 3 buckets: exacto, menor, mayor
5. Retorna el mejor de cada bucket

Esto da resultados muy buenos en milisegundos para cualquier tamaño de rango.

---

## Tecnologías usadas

- **Office.js** — API oficial de Microsoft para interactuar con Excel
- **HTML + CSS + JavaScript** — sin frameworks, fácil de mantener
- Funciona en **Excel Desktop** (Windows/Mac) y **Excel Online**
