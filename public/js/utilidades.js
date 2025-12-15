/* ============================================================
   📌 UTILIDADES GENERALES PARA TODAS LAS LIQUIDACIONES
   ============================================================ */

/* ------------------------------------------------------------
   🟦 1. LEER EXCEL (SheetJS)
------------------------------------------------------------ */
async function leerExcel(archivo) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const hoja = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(hoja, { defval: "" });
        resolve(json);
      } catch (error) {
        reject("Error al procesar el archivo Excel: " + error);
      }
    };

    reader.onerror = () => reject("No se pudo leer el archivo.");
    reader.readAsArrayBuffer(archivo);
  });
}

/* ------------------------------------------------------------
   🟩 2. LEER CSV
------------------------------------------------------------ */
async function leerCSV(archivo) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const lineas = e.target.result.split(/\r?\n/);
        const datos = lineas.map(l => l.split(","));
        resolve(datos);
      } catch (error) {
        reject("Error al leer CSV: " + error);
      }
    };

    reader.onerror = () => reject("No se pudo leer el archivo CSV.");
    reader.readAsText(archivo, "UTF-8");
  });
}

/* ------------------------------------------------------------
   🟧 3. LIMPIEZA DE IMPORTES
------------------------------------------------------------ */
function limpiarImporte(valor) {
  if (!valor) return 0;
  return parseFloat(
    valor.toString()
      .replace(/\./g, "")   // quita separador de miles
      .replace(",", ".")    // coma → punto
      .trim()
  ) || 0;
}

/* ------------------------------------------------------------
   🟪 4. FORMATEAR FECHAS
------------------------------------------------------------ */
function formatearFecha(fechaStr) {
  if (!fechaStr) return "";

  const partes = fechaStr.split(/[-\/]/);

  if (partes.length === 3) {
    let [d, m, a] = partes;

    // convierte AA → AAAA
    if (a.length === 2) a = "20" + a;

    return `${a}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }

  return fechaStr;
}

/* ------------------------------------------------------------
   🟦 5. DESCARGAR ARCHIVO (TXT / CSV)
------------------------------------------------------------ */
function descargarArchivo(nombre, contenido) {
  const blob = new Blob([contenido], { type: "text/plain" });
  const link = document.createElement("a");

  link.href = URL.createObjectURL(blob);
  link.download = nombre;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

/* ------------------------------------------------------------
   🟨 6. MOSTRAR TABLA DINÁMICA
------------------------------------------------------------ */
function mostrarTabla(contenedorId, datos) {
  const cont = document.getElementById(contenedorId);
  cont.innerHTML = "";

  if (!datos || datos.length === 0) {
    cont.innerHTML = "<p>No hay datos para mostrar.</p>";
    return;
  }

  const tabla = document.createElement("table");
  tabla.className = "tabla-datos";

  // encabezado
  const thead = document.createElement("thead");
  const filaHead = document.createElement("tr");

  Object.keys(datos[0]).forEach(col => {
    const th = document.createElement("th");
    th.textContent = col;
    filaHead.appendChild(th);
  });

  thead.appendChild(filaHead);
  tabla.appendChild(thead);

  // filas
  const tbody = document.createElement("tbody");

  datos.forEach(row => {
    const tr = document.createElement("tr");

    Object.values(row).forEach(val => {
      const td = document.createElement("td");
      td.textContent = val;
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  tabla.appendChild(tbody);
  cont.appendChild(tabla);
}

/* ------------------------------------------------------------
   🟥 7. MOSTRAR ERRORES
------------------------------------------------------------ */
function mostrarError(msg) {
  alert("❌ " + msg);
}

/* ------------------------------------------------------------
   🟫 8. LIMPIAR CONTENEDOR
------------------------------------------------------------ */
function limpiarContenedor(id) {
  document.getElementById(id).innerHTML = "";
}
