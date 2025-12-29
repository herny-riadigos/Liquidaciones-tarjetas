/* ============================================================
   🔵 Liquidación Nación — Versión Profesional (4 hojas)
   ============================================================ */

document.addEventListener("DOMContentLoaded", () => {
    const btn = document.getElementById("btnProcesarNacion");
    if (btn) btn.addEventListener("click", procesarNacion);
});


/* ============================================================
   🧩 FUNCIÓN PRINCIPAL
   ============================================================ */

function procesarNacion() {
    console.clear();

    let texto = document.getElementById("textoNacion").value.trim();
    if (!texto) {
        alert("Pegá una liquidación primero.");
        return;
    }

    texto = normalizarTexto(texto);

    // --- Extracción de datos ---
    const fechaPago = extraerFechaPago(texto);
    const totales = extraerTotales(texto);
    const tablas = extraerTablas(texto);
    const acreditacion = extraerAcreditacion(texto);

    // --- Crear libro Excel ---
    const wb = XLSX.utils.book_new();

    /* ------------------------------------------------------------
       HOJA 1 — RESUMEN
    ------------------------------------------------------------ */
    const hojaResumen = [];

    hojaResumen.push(["FECHA DE PAGO:", fechaPago || "NO IDENTIFICADA"]);
    hojaResumen.push([]);

    hojaResumen.push(["TOTALES"]);
    hojaResumen.push(["TOTAL PRESENTADO", totales.presentado || ""]);
    hojaResumen.push(["TOTAL DESCUENTO", totales.descuento || ""]);
    hojaResumen.push(["SALDO", totales.saldo || ""]);
    hojaResumen.push([]);

    if (acreditacion) {
        hojaResumen.push(["SE ACREDITÓ EN:", acreditacion.cuenta]);
        hojaResumen.push(["MONTO:", acreditacion.monto]);
    }

    const wsResumen = XLSX.utils.aoa_to_sheet(hojaResumen);
    aplicarEstiloTabla(wsResumen, hojaResumen.length);
    XLSX.utils.book_append_sheet(wb, wsResumen, "Resumen");


    /* ------------------------------------------------------------
       HOJA 2 — MOVIMIENTOS (TABLA 1)
    ------------------------------------------------------------ */

    const hojaMov = [];
    hojaMov.push(["MOVIMIENTOS (TABLA 1)"]);
    hojaMov.push([]);

    if (tablas[0]) tablas[0].forEach(fila => hojaMov.push(fila));
    else hojaMov.push(["No se detectaron movimientos."]);

    const wsMov = XLSX.utils.aoa_to_sheet(hojaMov);
    aplicarEstiloTabla(wsMov, hojaMov.length);
    XLSX.utils.book_append_sheet(wb, wsMov, "Movimientos");


    /* ------------------------------------------------------------
       HOJA 3 — DETALLES (TABLA 2)
    ------------------------------------------------------------ */

    const hojaDet = [];
    hojaDet.push(["DETALLES (TABLA 2)"]);
    hojaDet.push([]);

    if (tablas[1]) tablas[1].forEach(fila => hojaDet.push(fila));
    else hojaDet.push(["No se detectaron detalles."]);

    const wsDet = XLSX.utils.aoa_to_sheet(hojaDet);
    aplicarEstiloTabla(wsDet, hojaDet.length);
    XLSX.utils.book_append_sheet(wb, wsDet, "Detalles");


    /* ------------------------------------------------------------
       HOJA 4 — TEXTO ORIGINAL
    ------------------------------------------------------------ */

    const hojaTexto = texto.split("\n").map(x => [x]);
    const wsTexto = XLSX.utils.aoa_to_sheet(hojaTexto);
    XLSX.utils.book_append_sheet(wb, wsTexto, "Texto Original");


    /* ------------------------------------------------------------
       DESCARGAR ARCHIVO
    ------------------------------------------------------------ */

    const nombreSeguro = limpiarNombreHoja(fechaPago || "Liquidacion");
    XLSX.writeFile(wb, `Liquidacion_Nacion_${nombreSeguro}.xlsx`);

    alert("Liquidación procesada y descargada correctamente.");
}


/* ============================================================
   🔧 NORMALIZADOR DE TEXTO
   ============================================================ */

function normalizarTexto(txt) {
    return txt.replace(/\r/g, "")
              .replace(/\t/g, " ")
              .replace(/ +/g, " ")
              .trim();
}


/* ============================================================
   🔍 EXTRAER FECHA
   ============================================================ */

function extraerFechaPago(texto) {
    const r = texto.match(/fecha\s*de\s*pago[:\s]+(\d{1,2}\/\d{1,2}\/\d{4})/i);
    return r ? r[1] : "";
}


/* ============================================================
   🔍 EXTRAER TOTALES
   ============================================================ */

function extraerTotales(texto) {
    return {
        presentado: buscarMonto(texto, /total presentado/i),
        descuento: buscarMonto(texto, /total descuento/i),
        saldo: buscarMonto(texto, /saldo/i)
    };
}

function buscarMonto(texto, reg) {
    const lineas = texto.split("\n");
    for (let l of lineas) {
        if (reg.test(l)) {
            const m = l.match(/([\d.,]+)/);
            return m ? m[1] : "";
        }
    }
    return "";
}


/* ============================================================
   🔍 EXTRAER TABLAS (2 TABLAS PRINCIPALES)
   ============================================================ */

function extraerTablas(texto) {
    const lineas = texto.split("\n").map(l => l.trim()).filter(l => l !== "");
    let tablas = [];
    let tablaActual = [];

    lineas.forEach(l => {
        const cols = l.split(/\s+/);
        const filaEsTabla = /\d/.test(l) && cols.length >= 3 && cols.length <= 12;

        if (filaEsTabla) {
            tablaActual.push(cols);
        } else {
            if (tablaActual.length > 0) {
                tablas.push(tablaActual);
                tablaActual = [];
            }
        }
    });

    if (tablaActual.length > 0) tablas.push(tablaActual);

    return tablas;
}


/* ============================================================
   🔍 EXTRAER ACREDITACIÓN
   ============================================================ */

function extraerAcreditacion(texto) {
    const r = texto.match(/se acredit[oó] en[:\s]+([\d\s]+)\s*\$?([\d.,]+)/i);
    if (!r) return null;

    return {
        cuenta: r[1].trim(),
        monto: r[2].trim()
    };
}


/* ============================================================
   🔒 LIMPIAR NOMBRE DE HOJA
   ============================================================ */

function limpiarNombreHoja(nombre) {
    return nombre.replace(/[:\\\/\?\*\[\]]/g, "-");
}


/* ============================================================
   🎨 APLICAR ESTILOS (bordes + negrita en títulos)
   ============================================================ */

function aplicarEstiloTabla(ws, numFilas) {

    const range = XLSX.utils.decode_range(ws["!ref"]);

    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellRef = XLSX.utils.encode_cell({ r: R, c: C });

            if (!ws[cellRef]) continue;

            ws[cellRef].s = {
                border: {
                    top:    { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left:   { style: "thin", color: { rgb: "000000" } },
                    right:  { style: "thin", color: { rgb: "000000" } }
                },
                font: R === 0 ? { bold: true } : {}
            };
        }
    }
}
