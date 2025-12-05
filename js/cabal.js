/* ============================================================
   PARTE 1 - PARSER COMPLETO DE LIQUIDACIÓN CABAL
   ============================================================ */

/*
  Este parser devuelve un objeto con esta estructura:

  {
    encabezado: {
      fechaPago,
      liquidacionNro,
      cuenta
    },

    tit1: {
      cuadro1: [...],
      cuadro2: [...]
    },

    tit2: {
      totalVentas: { cantidad, total },
      cuadro: [...]
    },

    tit3: {
      cuadro1: [...],
      cuadro2: [...]
    },

    tit4: {
      totalVentas: { cantidad, total },
      cuadro: [...]
    },

    tit5: {
      totFechaPago: { fecha, cantidad, total },
      cuadro: [...],
      totalFinal: importeFinal
    }
  }
*/

function parsearLiquidacionCabal(textoOriginal) {
  // 1) LIMPIEZA DE BASURA
  let texto = textoOriginal
    .replace(/CONTINUA EN PAGINA SIGUIENTE.*?>>>/gi, "")
    .replace(/>>>>>> VIENE DE PAGINA ANTERIOR.*?\n/gi, "")
    .replace(/Banco Credicoop[\s\S]*?Argentina/gi, "")
    .replace(/ENCUADRES FISCALES[\s\S]*?INS/gi, "")
    .replace(/\n\s*\n/g, "\n")
    .trim();

  const lineas = texto.split("\n").map(l => l.trim());

  const datos = {
    encabezado: {},
    tit1: { cuadro1: [], cuadro2: [] },
    tit2: { totalVentas: {}, cuadro: [] },
    tit3: { cuadro1: [], cuadro2: [] },
    tit4: { totalVentas: {}, cuadro: [] },
    tit5: { totFechaPago: {}, cuadro: [], totalFinal: 0 }
  };

  /* ============================================================
     2) ENCABEZADO
     ============================================================ */

  lineas.forEach(l => {
    if (l.startsWith("FECHA DE PAGO"))
      datos.encabezado.fechaPago = l.split(":")[1].trim();

    if (l.startsWith("LIQUIDACION NRO"))
      datos.encabezado.liquidacionNro = l.split(":")[1].trim();

    if (l.startsWith("CUENTA P/ACREDITAR"))
      datos.encabezado.cuenta = l.split(":")[1].trim();
  });

  /* ============================================================
     3) TÍTULO 1 – VENTAS CORRESPONDIENTES A CABAL DEBITO
     ============================================================ */
  const idxTit1 = lineas.indexOf("VENTAS CORRESPONDIENTES A CABAL DEBITO");
  if (idxTit1 !== -1) {
    for (let i = idxTit1 + 1; i < lineas.length; i++) {
      let l = lineas[i];

      if (l.startsWith("08/") || /^\d{2}\/\d{2}\/\d{4}/.test(l)) {
        if (l.includes("*TOTAL*")) {
          // CUADRO 2
          const p = l.split(" ");
          datos.tit1.cuadro2.push({
            fecha: p[0],
            lote: p[1],
            terminal: p[2],
            cantidad: p[3],
            total: limpiarImporte(p[p.length - 1])
          });
        } else if (l.startsWith("CABAL DEBITO TOTAL")) {
          break;
        } else {
          // CUADRO 1
          const p = l.split(" ");
          const fecha = p[0];
          const nroCupon = p[1];
          const nroTarjeta = p[2];
          const cuota = p[3];
          const importe = limpiarImporte(p[4]);

          datos.tit1.cuadro1.push({
            fecha,
            nroCupon,
            nroTarjeta,
            cuota,
            importe
          });
        }
      }

      if (l.startsWith("CABAL DEBITO TOTAL")) break;
    }
  }

  /* ============================================================
     4) TÍTULO 2 – CABAL DEBITO (totales + cuadro)
     ============================================================ */
  lineas.forEach(l => {
    if (l.startsWith("CABAL DEBITO TOTAL DE VENTAS")) {
      const parts = l.split(" ");
      const cantidad = parseInt(parts[parts.length - 2]);
      const total = limpiarImporte(parts[parts.length - 1]);

      datos.tit2.totalVentas = { cantidad, total };
    }

    if (l.startsWith("ARANCEL DE DESCUENTO") &&
        !l.includes("TARJETA DE CREDITO")) {
      const parts = l.split(" ");
      const importe = limpiarImporte(parts.pop());
      const porc = extraerPorcentaje(l);
      const ref = extraerReferencia(l);

      datos.tit2.cuadro.push({
        concepto: "ARANCEL DE DESCUENTO",
        porc,
        referencia: ref,
        importe
      });
    }

    if (l.startsWith("IVA S/ARANCEL DE DESCUENTO")) {
      const importe = limpiarImporte(l.split(" ").pop());
      const porc = extraerPorcentaje(l);
      const ref = extraerReferencia(l);

      datos.tit2.cuadro.push({
        concepto: "IVA S/ARANCEL DE DESCUENTO",
        porc,
        referencia: ref,
        importe
      });
    }
  });

  /* ============================================================
     5) TÍTULO 3 – VENTAS TARJETA DE CREDITO
     ============================================================ */
  const idxTit3 = lineas.indexOf("VENTAS CORRESPONDIENTES A TARJETA DE CREDITO");
  if (idxTit3 !== -1) {
    for (let i = idxTit3 + 1; i < lineas.length; i++) {
      let l = lineas[i];

      if (l.startsWith("TARJETA DE CREDITO TOTAL")) break;

      if (/^\d{2}\/\d{2}\/\d{4}/.test(l)) {
        if (l.includes("*TOTAL*")) {
          const p = l.split(" ");
          datos.tit3.cuadro2.push({
            fecha: p[0],
            lote: p[1],
            terminal: p[2],
            cantidad: p[3],
            total: limpiarImporte(p[p.length - 1])
          });
        } else {
          const p = l.split(" ");
          datos.tit3.cuadro1.push({
            fecha: p[0],
            nroCupon: p[1],
            nroTarjeta: p[2],
            cuota: p[3],
            importe: limpiarImporte(p[4])
          });
        }
      }
    }
  }

  /* ============================================================
     6) TÍTULO 4 – TARJETA DE CREDITO (totales + cuadro)
     ============================================================ */

  lineas.forEach(l => {
    if (l.startsWith("TARJETA DE CREDITO TOTAL DE VENTAS")) {
      const p = l.split(" ");
      const cantidad = parseInt(p[p.length - 2]);
      const total = limpiarImporte(p[p.length - 1]);
      datos.tit4.totalVentas = { cantidad, total };
    }

    if (l.startsWith("ARANCEL DE DESCUENTO") &&
        !l.includes("CABAL DEBITO") &&
        !l.includes("TIT. FEC")) {
      const parts = l.split(" ");
      const importe = limpiarImporte(parts.pop());
      const porc = extraerPorcentaje(l);
      const ref = extraerReferencia(l);
      datos.tit4.cuadro.push({
        concepto: "ARANCEL DE DESCUENTO",
        porc,
        referencia: ref,
        importe
      });
    }

    if (l.startsWith("IVA S/ARANCEL DE DESCUENTO") &&
        !l.includes("CABAL DEBITO")) {
      const importe = limpiarImporte(l.split(" ").pop());
      const porc = extraerPorcentaje(l);
      const ref = extraerReferencia(l);
      datos.tit4.cuadro.push({
        concepto: "IVA S/ARANCEL DE DESCUENTO",
        porc,
        referencia: ref,
        importe
      });
    }
  });

  /* ============================================================
     7) TÍTULO 5 – TOT FEC PAGO (final)
     ============================================================ */

  lineas.forEach(l => {
    if (l.startsWith("TOT. FEC. PAGO")) {
      const parts = l.split(" ");

      const fecha = parts[3];
      const cantidad = parseInt(parts[parts.length - 2]);
      const total = limpiarImporte(parts[parts.length - 1]);

      datos.tit5.totFechaPago = { fecha, cantidad, total };
    }

    const conceptosFinales = [
      "ARANCEL DE DESCUENTO",
      "IVA S/ARANCEL DE DESCUENTO",
      "IVA S/ARANCEL + COSTO FINANCIERO",
      "COSTO FINANCIERO",
      "NETO A LIQUIDAR POR VENTAS",
      "PERCEPCION DE IVA RG 333",
      "PERCEPCION DE IIBB",
      "RETENCION IIBB SIRTAC"
    ];

    for (const c of conceptosFinales) {
      if (l.startsWith(c)) {
        const partes = l.split(" ");

        const importe = limpiarImporte(partes.pop());
        const ref = extraerReferencia(l);
        const porc = extraerPorcentaje(l);

        datos.tit5.cuadro.push({
          concepto: c,
          porc,
          referencia: ref,
          importe
        });
      }
    }

    if (l.startsWith("IMPORTE NETO FINAL")) {
      datos.tit5.totalFinal = limpiarImporte(l.split(" ").pop());
    }
  });

  return datos;
}

/* ============================================================
   FUNCIONES AUXILIARES
   ============================================================ */

function extraerPorcentaje(l) {
  const m = l.match(/(\d+,\d+)%/);
  return m ? m[1] : "";
}

function extraerReferencia(l) {
  const m = l.match(/(\d+,\d+)\s+\d+,\d+−?$/);
  return m ? m[1] : "";
}

/* ============================================================
   PARTE 2 - GENERADOR DE EXCEL
   ============================================================ */

function generarExcelCabal(datos, nombreArchivo) {
  const wb = XLSX.utils.book_new();
  const ws = {};

  let fila = 1; // Va aumentando a medida que agregamos contenido

  function setCell(col, row, value) {
    const cell = col + row;
    ws[cell] = { v: value };
  }

  function escribirFila(titulo, arrColumnas, arrDatos) {
    // Título combinado
    setCell("A", fila, titulo);
    fila++;

    // Encabezados
    arrColumnas.forEach((col, idx) => {
      setCell(String.fromCharCode(65 + idx), fila, col);
    });
    fila++;

    // Datos
    arrDatos.forEach(reg => {
      let colIdx = 0;
      for (let k in reg) {
        setCell(String.fromCharCode(65 + colIdx), fila, reg[k]);
        colIdx++;
      }
      fila++;
    });

    // Separador visual
    fila++;
  }

  /* ============================================================
       BLOQUE ENCABEZADO
     ============================================================ */
  setCell("A", fila, "ENCABEZADO");
  fila++;

  setCell("A", fila, "FECHA DE PAGO");
  setCell("B", fila, datos.encabezado.fechaPago);
  fila++;

  setCell("A", fila, "NRO LIQUIDACION");
  setCell("B", fila, datos.encabezado.liquidacionNro);
  fila++;

  setCell("A", fila, "CUENTA");
  setCell("B", fila, datos.encabezado.cuenta);
  fila += 2; // separador

  /* ============================================================
       TÍTULO 1 – CABAL DÉBITO (CUADRO 1 + CUADRO 2)
     ============================================================ */

  // CUADRO 1
  escribirFila(
    "VENTAS CORRESPONDIENTES A CABAL DEBITO — CUADRO 1",
    ["FECHA COMPRA", "NRO CUPON", "NRO TARJETA", "CUOTA", "IMPORTE TOTAL"],
    datos.tit1.cuadro1
  );

  // CUADRO 2
  escribirFila(
    "VENTAS CORRESPONDIENTES A CABAL DEBITO — CUADRO 2",
    ["FECHA COMPRA", "NRO LOTE", "NRO TERMINAL", "CANTIDAD CUPONES", "TOTAL"],
    datos.tit1.cuadro2
  );

  /* ============================================================
       TÍTULO 2 – CABAL DEBITO (TOTAL + CUADRO)
     ============================================================ */

  // Total de ventas
  escribirFila(
    "CABAL DEBITO — TOTAL DE VENTAS",
    ["CANTIDAD", "TOTAL"],
    [datos.tit2.totalVentas]
  );

  // Cuadro (arancel + iva)
  escribirFila(
    "CABAL DEBITO — CUADRO",
    ["CONCEPTO", "%", "REFERENCIA", "IMPORTE TOTAL"],
    datos.tit2.cuadro
  );

  /* ============================================================
       TÍTULO 3 – TARJETA CRÉDITO (CUADRO 1 + 2)
     ============================================================ */

  escribirFila(
    "VENTAS CORRESPONDIENTES A TARJETA DE CREDITO — CUADRO 1",
    ["FECHA COMPRA", "NRO CUPON", "NRO TARJETA", "CUOTA", "IMPORTE TOTAL"],
    datos.tit3.cuadro1
  );

  escribirFila(
    "VENTAS CORRESPONDIENTES A TARJETA DE CREDITO — CUADRO 2",
    ["FECHA COMPRA", "NRO LOTE", "NRO TERMINAL", "CANTIDAD CUPONES", "TOTAL"],
    datos.tit3.cuadro2
  );

  /* ============================================================
       TÍTULO 4 – TARJETA DE CRÉDITO (TOTAL + CUADRO)
     ============================================================ */

  escribirFila(
    "TARJETA DE CREDITO — TOTAL DE VENTAS",
    ["CANTIDAD", "TOTAL"],
    [datos.tit4.totalVentas]
  );

  escribirFila(
    "TARJETA DE CREDITO — CUADRO",
    ["CONCEPTO", "%", "REFERENCIA", "IMPORTE TOTAL"],
    datos.tit4.cuadro
  );

  /* ============================================================
       TÍTULO 5 – TOT FEC PAGO
     ============================================================ */

  // Total fecha pago
  escribirFila(
    "TOT FEC PAGO — TOTAL",
    ["FECHA", "CANTIDAD", "TOTAL"],
    [datos.tit5.totFechaPago]
  );

  // Cuadro final (retenciones, percepciones, neto, arancel)
  escribirFila(
    "TOT FEC PAGO — CUADRO FINAL",
    ["CONCEPTO", "%", "REFERENCIA", "IMPORTE TOTAL"],
    datos.tit5.cuadro
  );

  // Total final (último renglón)
  setCell("A", fila, "IMPORTE NETO FINAL");
  setCell("B", fila, datos.tit5.totalFinal);

  ws["!ref"] = `A1:Z${fila}`;
  XLSX.utils.book_append_sheet(wb, ws, "LIQ CABAL");
  XLSX.writeFile(wb, nombreArchivo + ".xlsx");
}

/* ============================================================
   PARTE 3 - INTEGRACIÓN ADAPTADA A TU cabal.html
   ============================================================ */

let liquidacionesProcesadas = [];

// Esperar que el HTML esté cargado
document.addEventListener("DOMContentLoaded", () => {

  const btnProcesar = document.getElementById("btnProcesar");
  const btnTotalizador = document.getElementById("btnTotalizador");

  // Si no existen, mostrar error en consola
  if (!btnProcesar) {
    console.error("ERROR: No existe el botón #btnProcesar en el HTML");
    return;
  }
  if (!btnTotalizador) {
    console.error("ERROR: No existe el botón #btnTotalizador en el HTML");
    return;
  }

  btnProcesar.addEventListener("click", procesarLiquidacionCabal);
  btnTotalizador.addEventListener("click", armarTotalizadorCabal);
});

/* ============================================================
   MANEJA EL CLICK EN "PROCESAR LIQUIDACIÓN"
   ============================================================ */

function procesarLiquidacionCabal() {
  const textarea = document.getElementById("textoCabal");
  const texto = textarea.value.trim();

  if (texto === "") {
    alert("Pegá una liquidación antes de procesar.");
    return;
  }

  // 1) Parseo
  const datos = parsearLiquidacionCabal(texto);

  // 2) Guardar internamente
  liquidacionesProcesadas.push(datos);

  // 3) Generar Excel individual
  const nombreArchivo = `Liquidacion_${datos.encabezado.liquidacionNro}`;
  generarExcelCabal(datos, nombreArchivo);

  // 4) Agregar a la lista lateral
  agregarALaListaCabal(nombreArchivo);

  // 5) Vaciar textarea para pegar otra liq
  textarea.value = "";
}

/* ============================================================
   AGREGA A LA LISTA LATERAL
   ============================================================ */

function agregarALaListaCabal(nombre) {
  const ul = document.getElementById("listaProcesadas");
  const li = document.createElement("li");
  li.textContent = nombre;
  ul.appendChild(li);
}

/* ============================================================
   BOTÓN "ARMAR TOTALIZADOR"
   ============================================================ */

function armarTotalizadorCabal() {
  if (liquidacionesProcesadas.length === 0) {
    alert("Procesá al menos una liquidación primero.");
    return;
  }

  alert("El Totalizador se implementará luego. Ya tengo todos los datos preparados.");
}
