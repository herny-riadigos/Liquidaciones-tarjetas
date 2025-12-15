/* ============================================================
   PARTE 1 - PARSER COMPLETO DE LIQUIDACIÓN CABAL - VERSIÓN CORREGIDA
   ============================================================ */

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
     2) ENCABEZADO - Versión mejorada para líneas combinadas
     ============================================================ */

  // Unir primeras líneas para buscar el encabezado
  let textoEncabezado = lineas.slice(0, 3).join(" ");
  
  // Buscar FECHA DE PAGO con regex
  const regexFecha = /FECHA\s+DE\s+PAGO\s*:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i;
  const matchFecha = textoEncabezado.match(regexFecha);
  if (matchFecha) {
    datos.encabezado.fechaPago = matchFecha[1];
  }

  // Buscar LIQUIDACION NRO
  const regexLiquidacion = /LIQUIDACION\s+NRO\.?\s*:?\s*(\d+)/i;
  const matchLiquidacion = textoEncabezado.match(regexLiquidacion);
  if (matchLiquidacion) {
    datos.encabezado.liquidacionNro = matchLiquidacion[1];
  }

  // Buscar CUENTA P/ACREDITAR VENTAS
  const regexCuenta = /CUENTA\s+P\/ACREDITAR\s+VENTAS\s*:?\s*([^:]+?(?=\s*(?:ENCUADRES\s+FISCALES|$)))/i;
  const matchCuenta = textoEncabezado.match(regexCuenta);
  if (matchCuenta) {
    datos.encabezado.cuenta = matchCuenta[1].trim();
  }

  // Si no encontró con regex, probar el método original línea por línea
  if (!datos.encabezado.fechaPago || !datos.encabezado.liquidacionNro) {
    lineas.forEach(l => {
      if (l.toUpperCase().includes("FECHA DE PAGO")) {
        datos.encabezado.fechaPago = extraerValorMejorado(l, "FECHA DE PAGO");
      }
      if (l.toUpperCase().includes("LIQUIDACION NRO")) {
        datos.encabezado.liquidacionNro = extraerValorMejorado(l, "LIQUIDACION NRO");
      }
      if (l.toUpperCase().includes("CUENTA P/ACREDITAR VENTAS")) {
        datos.encabezado.cuenta = extraerValorMejorado(l, "CUENTA P/ACREDITAR VENTAS");
      }
    });
  }

  /* ============================================================
     3) TÍTULO 1 – VENTAS CORRESPONDIENTES A CABAL DEBITO
     ============================================================ */
  const idxTit1 = lineas.findIndex(l => 
    l.toUpperCase().includes("VENTAS CORRESPONDIENTES A CABAL DEBITO")
  );
  
  if (idxTit1 !== -1) {
    for (let i = idxTit1 + 1; i < lineas.length; i++) {
      let l = lineas[i];

      // Detener si encontramos el siguiente título
      if (l.toUpperCase().includes("CABAL DEBITO TOTAL") ||
          l.toUpperCase().includes("VENTAS CORRESPONDIENTES A TARJETA DE CREDITO")) {
        break;
      }

      // Procesar líneas con formato de fecha
      if (l.startsWith("08/") || /^\d{2}\/\d{2}\/\d{4}/.test(l)) {
        if (l.includes("*TOTAL*")) {
          // CUADRO 2 (Total por terminal)
          const partes = l.split(" ").filter(p => p !== "");
          if (partes.length >= 5) {
            datos.tit1.cuadro2.push({
              fecha: partes[0],
              lote: partes[1],
              terminal: partes[2],
              cantidad: partes[3],
              total: limpiarImporte(partes[partes.length - 1])
            });
          }
        } else {
          // CUADRO 1 (Detalle de ventas)
          const partes = l.split(" ").filter(p => p !== "");
          if (partes.length >= 5) {
            datos.tit1.cuadro1.push({
              fecha: partes[0],
              nroCupon: partes[1],
              nroTarjeta: partes[2],
              cuota: partes[3],
              importe: limpiarImporte(partes[4])
            });
          }
        }
      }
    }
  }

  /* ============================================================
     4) TÍTULO 2 – CABAL DEBITO (totales + cuadro)
     ============================================================ */
  lineas.forEach(l => {
    if (l.toUpperCase().startsWith("CABAL DEBITO TOTAL DE VENTAS")) {
      const partes = l.split(" ").filter(p => p !== "");
      const cantidad = parseInt(partes[partes.length - 2]) || 0;
      const total = limpiarImporte(partes[partes.length - 1]);
      datos.tit2.totalVentas = { cantidad, total };
    }

    if (l.toUpperCase().includes("ARANCEL DE DESCUENTO") &&
        l.toUpperCase().includes("CABAL") &&
        !l.toUpperCase().includes("CREDITO")) {
      const partes = l.split(" ").filter(p => p !== "");
      const importe = limpiarImporte(partes[partes.length - 1]);
      const porc = extraerPorcentaje(l);
      const ref = extraerReferencia(l);

      datos.tit2.cuadro.push({
        concepto: "ARANCEL DE DESCUENTO (DÉBITO)",
        porc,
        referencia: ref,
        importe
      });
    }

    if (l.toUpperCase().includes("IVA S/ARANCEL DE DESCUENTO") &&
        l.toUpperCase().includes("CABAL") &&
        !l.toUpperCase().includes("CREDITO")) {
      const partes = l.split(" ").filter(p => p !== "");
      const importe = limpiarImporte(partes[partes.length - 1]);
      const porc = extraerPorcentaje(l);
      const ref = extraerReferencia(l);

      datos.tit2.cuadro.push({
        concepto: "IVA S/ARANCEL DE DESCUENTO (DÉBITO)",
        porc,
        referencia: ref,
        importe
      });
    }
  });

  /* ============================================================
     5) TÍTULO 3 – VENTAS TARJETA DE CREDITO
     ============================================================ */
  const idxTit3 = lineas.findIndex(l => 
    l.toUpperCase().includes("VENTAS CORRESPONDIENTES A TARJETA DE CREDITO")
  );
  
  if (idxTit3 !== -1) {
    for (let i = idxTit3 + 1; i < lineas.length; i++) {
      let l = lineas[i];

      if (l.toUpperCase().includes("TARJETA DE CREDITO TOTAL")) break;

      if (/^\d{2}\/\d{2}\/\d{4}/.test(l)) {
        if (l.includes("*TOTAL*")) {
          const partes = l.split(" ").filter(p => p !== "");
          if (partes.length >= 5) {
            datos.tit3.cuadro2.push({
              fecha: partes[0],
              lote: partes[1],
              terminal: partes[2],
              cantidad: partes[3],
              total: limpiarImporte(partes[partes.length - 1])
            });
          }
        } else {
          const partes = l.split(" ").filter(p => p !== "");
          if (partes.length >= 5) {
            datos.tit3.cuadro1.push({
              fecha: partes[0],
              nroCupon: partes[1],
              nroTarjeta: partes[2],
              cuota: partes[3],
              importe: limpiarImporte(partes[4])
            });
          }
        }
      }
    }
  }

  /* ============================================================
     6) TÍTULO 4 – TARJETA DE CREDITO (totales + cuadro)
     ============================================================ */

  lineas.forEach(l => {
    if (l.toUpperCase().startsWith("TARJETA DE CREDITO TOTAL DE VENTAS")) {
      const partes = l.split(" ").filter(p => p !== "");
      const cantidad = parseInt(partes[partes.length - 2]) || 0;
      const total = limpiarImporte(partes[partes.length - 1]);
      datos.tit4.totalVentas = { cantidad, total };
    }

    if (l.toUpperCase().includes("ARANCEL DE DESCUENTO") &&
        l.toUpperCase().includes("CREDITO") &&
        !l.toUpperCase().includes("DEBITO")) {
      const partes = l.split(" ").filter(p => p !== "");
      const importe = limpiarImporte(partes[partes.length - 1]);
      const porc = extraerPorcentaje(l);
      const ref = extraerReferencia(l);
      
      datos.tit4.cuadro.push({
        concepto: "ARANCEL DE DESCUENTO (CRÉDITO)",
        porc,
        referencia: ref,
        importe
      });
    }

    if (l.toUpperCase().includes("IVA S/ARANCEL DE DESCUENTO") &&
        l.toUpperCase().includes("CREDITO") &&
        !l.toUpperCase().includes("DEBITO")) {
      const partes = l.split(" ").filter(p => p !== "");
      const importe = limpiarImporte(partes[partes.length - 1]);
      const porc = extraerPorcentaje(l);
      const ref = extraerReferencia(l);
      
      datos.tit4.cuadro.push({
        concepto: "IVA S/ARANCEL DE DESCUENTO (CRÉDITO)",
        porc,
        referencia: ref,
        importe
      });
    }
  });

  /* ============================================================
     7) TÍTULO 5 – TOT FEC PAGO (FINAL) - VERSIÓN CORREGIDA
     ============================================================ */
  
  // Buscar el inicio de la sección final (TOT. FEC. PAGO)
  const inicioSeccionFinal = lineas.findIndex(l => 
    l.toUpperCase().startsWith("TOT. FEC. PAGO")
  );
  
  if (inicioSeccionFinal !== -1) {
    // Procesar TOT. FEC. PAGO
    const lineaTotFecPago = lineas[inicioSeccionFinal];
    const partesTotFec = lineaTotFecPago.split(" ").filter(p => p !== "");
    if (partesTotFec.length >= 6) {
      datos.tit5.totFechaPago = {
        fecha: partesTotFec[3],
        cantidad: parseInt(partesTotFec[partesTotFec.length - 2]) || 0,
        total: limpiarImporte(partesTotFec[partesTotFec.length - 1])
      };
    }
    
    // Procesar desde TOT. FEC. PAGO hasta IMPORTE NETO FINAL
    // Los conceptos finales comienzan después de TOT. FEC. PAGO
    for (let i = inicioSeccionFinal + 1; i < lineas.length; i++) {
      const l = lineas[i];
      
      // Detener cuando encontramos IMPORTE NETO FINAL
      if (l.toUpperCase().includes("IMPORTE NETO FINAL")) {
        // Extraer el importe neto final
        const partes = l.split(" ").filter(p => p !== "");
        datos.tit5.totalFinal = limpiarImporte(partes[partes.length - 1]);
        break;
      }
      
      // Procesar conceptos finales (solo los que están en esta sección)
      if (l.includes("ARANCEL DE DESCUENTO") && 
          !l.includes("IVA") && 
          !l.includes("CABAL") && 
          !l.includes("CREDITO")) {
        // ARANCEL DE DESCUENTO (total general)
        const importe = limpiarImporte(l.split(" ").pop());
        datos.tit5.cuadro.push({
          concepto: "ARANCEL DE DESCUENTO (TOTAL)",
          porc: "",
          referencia: "",
          importe
        });
      }
      
      if (l.includes("IVA S/ARANCEL + COSTO FINANCIERO")) {
        const importe = limpiarImporte(l.split(" ").pop());
        datos.tit5.cuadro.push({
          concepto: "IVA S/ARANCEL + COSTO FINANCIERO",
          porc: "",
          referencia: "",
          importe
        });
      }
      
      if (l.includes("NETO A LIQUIDAR POR VENTAS")) {
        const importe = limpiarImporte(l.split(" ").pop());
        datos.tit5.cuadro.push({
          concepto: "NETO A LIQUIDAR POR VENTAS",
          porc: "",
          referencia: "",
          importe
        });
      }
      
      if (l.includes("PERCEPCION DE IVA RG 333")) {
        const importe = limpiarImporte(l.split(" ").pop());
        const porc = extraerPorcentaje(l);
        const ref = extraerReferencia(l);
        datos.tit5.cuadro.push({
          concepto: "PERCEPCION DE IVA RG 333",
          porc,
          referencia: ref,
          importe
        });
      }
      
      if (l.includes("PERCEPCION DE IIBB")) {
        const importe = limpiarImporte(l.split(" ").pop());
        const porc = extraerPorcentaje(l);
        const ref = extraerReferencia(l);
        datos.tit5.cuadro.push({
          concepto: "PERCEPCION DE IIBB",
          porc,
          referencia: ref,
          importe
        });
      }
      
      if (l.includes("RETENCION IIBB SIRTAC")) {
        const importe = limpiarImporte(l.split(" ").pop());
        const porc = extraerPorcentaje(l);
        const ref = extraerReferencia(l);
        datos.tit5.cuadro.push({
          concepto: "RETENCION IIBB SIRTAC",
          porc,
          referencia: ref,
          importe
        });
      }
    }
  }

  return datos;
}

/* ============================================================
   FUNCIONES AUXILIARES
   ============================================================ */

function limpiarImporte(texto) {
  if (!texto) return 0;
  // Reemplazar comas por puntos y eliminar caracteres no numéricos
  const limpio = texto.replace(/\./g, "").replace(/,/g, ".").replace(/[^\d.-]/g, "");
  const numero = parseFloat(limpio);
  return isNaN(numero) ? 0 : numero;
}

function extraerPorcentaje(l) {
  const m = l.match(/(\d+,\d+)%/);
  return m ? m[1].replace(',', '.') : "";
}

function extraerReferencia(l) {
  // Busca el primer número con decimales (puede ser con coma o punto)
  const m = l.match(/(\d+[.,]\d+)\s+\d+[.,]\d+/);
  return m ? m[1].replace(',', '.') : "";
}

function extraerValorMejorado(linea, clave) {
  // Busca la clave y extrae todo lo que viene después
  const indice = linea.toUpperCase().indexOf(clave.toUpperCase());
  if (indice === -1) return "";
  
  // Tomar la parte después de la clave
  const parte = linea.substring(indice + clave.length);
  // Buscar dos puntos y tomar lo que sigue
  const indiceDosPuntos = parte.indexOf(":");
  if (indiceDosPuntos !== -1) {
    const valor = parte.substring(indiceDosPuntos + 1).trim();
    // Limpiar si hay basura después
    const indiceBasura = valor.toUpperCase().indexOf("ENCUADRES FISCALES");
    if (indiceBasura !== -1) {
      return valor.substring(0, indiceBasura).trim();
    }
    return valor;
  }
  return parte.trim();
}

/* ============================================================
   PARTE 2 - GENERADOR DE EXCEL CON FORMATO DE MONEDA
   ============================================================ */

function generarExcelCabal(datos, nombreArchivo) {
  // Crear un nuevo libro de trabajo
  const wb = XLSX.utils.book_new();
  
  // Crearemos la hoja manualmente para controlar el formato
  const ws = {};
  
  // Configuración de ancho de columnas
  const wscols = [
    { wch: 30 }, // Columna A
    { wch: 10 }, // Columna B  
    { wch: 15 }, // Columna C
    { wch: 20 }, // Columna D
    { wch: 20 }  // Columna E
  ];
  
  // Función para agregar celda con formato
  function agregarCelda(col, row, valor, esNumero = false, esMoneda = false) {
    const celda = XLSX.utils.encode_cell({c: col, r: row});
    
    if (esNumero) {
      // Convertir a número
      const numValor = typeof valor === 'string' ? parseFloat(valor) : valor;
      ws[celda] = { 
        v: numValor,
        t: 'n'
      };
      
      if (esMoneda) {
        // Formato de moneda argentina
        ws[celda].z = '"$"#,##0.00;[Red]"$"#,##0.00';
      } else {
        // Formato numérico estándar
        ws[celda].z = '#,##0.00';
      }
    } else {
      // Texto normal
      ws[celda] = { v: valor, t: 's' };
    }
  }
  
  // Función para agregar fila
  function agregarFila(fila, valores) {
    valores.forEach((valor, col) => {
      // Determinar si es número
      const esNumero = typeof valor === 'number';
      const esMoneda = esNumero; // Todos los números serán moneda por ahora
      agregarCelda(col, fila, valor, esNumero, esMoneda);
    });
  }
  
  let filaActual = 0;
  
  // ENCABEZADO
  agregarFila(filaActual++, ["ENCABEZADO DE LIQUIDACIÓN"]);
  agregarFila(filaActual++, ["FECHA DE PAGO:", datos.encabezado.fechaPago]);
  agregarFila(filaActual++, ["NRO. LIQUIDACIÓN:", datos.encabezado.liquidacionNro]);
  agregarFila(filaActual++, ["CUENTA:", datos.encabezado.cuenta]);
  filaActual++; // Línea vacía
  
  // TÍTULO 1: VENTAS CABAL DÉBITO - CUADRO 1
  if (datos.tit1.cuadro1.length > 0) {
    agregarFila(filaActual++, ["VENTAS CORRESPONDIENTES A CABAL DÉBITO - DETALLE"]);
    agregarFila(filaActual++, ["FECHA", "NRO CUPÓN", "NRO TARJETA", "CUOTA", "IMPORTE"]);
    
    datos.tit1.cuadro1.forEach(item => {
      agregarFila(filaActual++, [
        item.fecha,
        item.nroCupon,
        item.nroTarjeta,
        item.cuota,
        item.importe
      ]);
    });
    filaActual++;
  }
  
  // TÍTULO 1: VENTAS CABAL DÉBITO - CUADRO 2
  if (datos.tit1.cuadro2.length > 0) {
    agregarFila(filaActual++, ["VENTAS CORRESPONDIENTES A CABAL DÉBITO - TOTALES POR TERMINAL"]);
    agregarFila(filaActual++, ["FECHA", "LOTE", "TERMINAL", "CANTIDAD", "TOTAL"]);
    
    datos.tit1.cuadro2.forEach(item => {
      agregarFila(filaActual++, [
        item.fecha,
        item.lote,
        item.terminal,
        item.cantidad,
        item.total
      ]);
    });
    filaActual++;
  }
  
  // TÍTULO 2: CABAL DÉBITO TOTALES
  agregarFila(filaActual++, ["CABAL DÉBITO - RESUMEN"]);
  agregarFila(filaActual++, ["TOTAL VENTAS:", `Cantidad: ${datos.tit2.totalVentas.cantidad || 0}`, `Importe:`, "", datos.tit2.totalVentas.total || 0]);
  
  if (datos.tit2.cuadro.length > 0) {
    filaActual++;
    agregarFila(filaActual++, ["CONCEPTO", "%", "REFERENCIA", "", "IMPORTE"]);
    
    datos.tit2.cuadro.forEach(item => {
      agregarFila(filaActual++, [
        item.concepto,
        item.porc,
        item.referencia,
        "",
        item.importe
      ]);
    });
  }
  filaActual++;
  
  // TÍTULO 3: VENTAS TARJETA CRÉDITO - CUADRO 1
  if (datos.tit3.cuadro1.length > 0) {
    agregarFila(filaActual++, ["VENTAS TARJETA DE CRÉDITO - DETALLE"]);
    agregarFila(filaActual++, ["FECHA", "NRO CUPÓN", "NRO TARJETA", "CUOTA", "IMPORTE"]);
    
    datos.tit3.cuadro1.forEach(item => {
      agregarFila(filaActual++, [
        item.fecha,
        item.nroCupon,
        item.nroTarjeta,
        item.cuota,
        item.importe
      ]);
    });
    filaActual++;
  }
  
  // TÍTULO 3: VENTAS TARJETA CRÉDITO - CUADRO 2
  if (datos.tit3.cuadro2.length > 0) {
    agregarFila(filaActual++, ["VENTAS TARJETA DE CRÉDITO - TOTALES POR TERMINAL"]);
    agregarFila(filaActual++, ["FECHA", "LOTE", "TERMINAL", "CANTIDAD", "TOTAL"]);
    
    datos.tit3.cuadro2.forEach(item => {
      agregarFila(filaActual++, [
        item.fecha,
        item.lote,
        item.terminal,
        item.cantidad,
        item.total
      ]);
    });
    filaActual++;
  }
  
  // TÍTULO 4: TARJETA CRÉDITO TOTALES
  agregarFila(filaActual++, ["TARJETA DE CRÉDITO - RESUMEN"]);
  agregarFila(filaActual++, ["TOTAL VENTAS:", `Cantidad: ${datos.tit4.totalVentas.cantidad || 0}`, `Importe:`, "", datos.tit4.totalVentas.total || 0]);
  
  if (datos.tit4.cuadro.length > 0) {
    filaActual++;
    agregarFila(filaActual++, ["CONCEPTO", "%", "REFERENCIA", "", "IMPORTE"]);
    
    datos.tit4.cuadro.forEach(item => {
      agregarFila(filaActual++, [
        item.concepto,
        item.porc,
        item.referencia,
        "",
        item.importe
      ]);
    });
  }
  filaActual++;
  
  // TÍTULO 5: TOT FEC PAGO
  if (datos.tit5.totFechaPago.fecha) {
    agregarFila(filaActual++, ["TOTAL FECHA DE PAGO"]);
    agregarFila(filaActual++, ["FECHA:", datos.tit5.totFechaPago.fecha]);
    agregarFila(filaActual++, ["CANTIDAD:", datos.tit5.totFechaPago.cantidad]);
    agregarFila(filaActual++, ["TOTAL:", "", "", "", datos.tit5.totFechaPago.total]);
    filaActual++;
  }
  
  // TÍTULO 5: CUADRO FINAL (SOLO CONCEPTOS FINALES)
  if (datos.tit5.cuadro.length > 0) {
    agregarFila(filaActual++, ["CONCEPTOS FINALES"]);
    agregarFila(filaActual++, ["CONCEPTO", "%", "REFERENCIA", "", "IMPORTE"]);
    
    datos.tit5.cuadro.forEach(item => {
      agregarFila(filaActual++, [
        item.concepto,
        item.porc,
        item.referencia,
        "",
        item.importe
      ]);
    });
    filaActual++;
  }
  
  // IMPORTE NETO FINAL
  if (datos.tit5.totalFinal) {
    agregarFila(filaActual++, ["IMPORTE NETO FINAL A LIQUIDAR"]);
    agregarFila(filaActual++, ["IMPORTE:", "", "", "", datos.tit5.totalFinal]);
  }
  
  // Definir el rango de la hoja
  const range = {s: {c: 0, r: 0}, e: {c: 4, r: filaActual}};
  ws['!ref'] = XLSX.utils.encode_range(range);
  ws['!cols'] = wscols;
  
  // Agregar la hoja al libro
  XLSX.utils.book_append_sheet(wb, ws, "Liquidación");
  
  // Descargar el archivo
  const nombreFinal = `${nombreArchivo || "liquidacion_cabal"}.xlsx`;
  XLSX.writeFile(wb, nombreFinal);
  
  return nombreFinal;
}


/* ============================================================
   PARTE 3 - INTEGRACIÓN CON LA INTERFAZ
   ============================================================ */

let liquidacionesProcesadas = [];

// Esperar que el HTML esté cargado
document.addEventListener("DOMContentLoaded", () => {
  const btnProcesar = document.getElementById("btnProcesar");
  const btnTotalizador = document.getElementById("btnTotalizador");
  const btnLimpiar = document.getElementById("btnLimpiar");

  if (btnProcesar) {
    btnProcesar.addEventListener("click", procesarLiquidacionCabal);
  }
  
  if (btnTotalizador) {
    btnTotalizador.addEventListener("click", armarTotalizadorCabal);
  }
  
  if (btnLimpiar) {
    btnLimpiar.addEventListener("click", () => {
      document.getElementById("textoCabal").value = "";
      document.getElementById("listaProcesadas").innerHTML = "";
      liquidacionesProcesadas = [];
    });
  }
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

  try {
    // 1) Parsear los datos
    const datos = parsearLiquidacionCabal(texto);
    
    console.log("Datos parseados:", datos); // Para depuración
    
    // 2) Verificar que se extrajo información
    if (!datos.encabezado.fechaPago && !datos.encabezado.liquidacionNro) {
      alert("No se pudo extraer información del encabezado. Verifica el formato.");
      return;
    }
    
    // 3) Guardar internamente
    liquidacionesProcesadas.push(datos);
    
    // 4) Generar nombre de archivo
    const numLiquidacion = datos.encabezado.liquidacionNro || "SIN_NUMERO";
    const fecha = datos.encabezado.fechaPago ? datos.encabezado.fechaPago.replace(/\//g, "-") : "";
    const nombreArchivo = `Liquidacion_CABAL_${numLiquidacion}_${fecha}`;
    
    // 5) Generar y descargar Excel
    generarExcelCabal(datos, nombreArchivo);
    
    // 6) Agregar a la lista lateral
    agregarALaListaCabal(nombreArchivo, datos.encabezado);
    
    // 7) Mostrar confirmación
    mostrarConfirmacion(datos.encabezado);
    
    // 8) Opcional: vaciar textarea
    // textarea.value = "";
    
  } catch (error) {
    console.error("Error procesando liquidación:", error);
    alert(`Error al procesar la liquidación: ${error.message}`);
  }
}

/* ============================================================
   AGREGA A LA LISTA LATERAL
   ============================================================ */

function agregarALaListaCabal(nombreArchivo, encabezado) {
  const ul = document.getElementById("listaProcesadas");
  if (!ul) return;
  
  const li = document.createElement("li");
  li.className = "list-group-item d-flex justify-content-between align-items-center";
  
  const info = document.createElement("span");
  info.textContent = `${encabezado.liquidacionNro || "Sin N°"} - ${encabezado.fechaPago || "Sin fecha"}`;
  
  const badge = document.createElement("span");
  badge.className = "badge bg-success rounded-pill";
  badge.textContent = "✓";
  
  li.appendChild(info);
  li.appendChild(badge);
  ul.appendChild(li);
}

/* ============================================================
   MUESTRA CONFIRMACIÓN
   ============================================================ */

function mostrarConfirmacion(encabezado) {
  const confirmacion = document.createElement("div");
  confirmacion.className = "alert alert-success alert-dismissible fade show mt-3";
  confirmacion.innerHTML = `
    <strong>✓ Liquidación procesada correctamente</strong>
    <p>N° Liquidación: ${encabezado.liquidacionNro || "No identificado"}<br>
       Fecha: ${encabezado.fechaPago || "No especificada"}<br>
       <small>El archivo Excel se está descargando automáticamente.</small>
    </p>
    <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
  `;
  
  const contenedor = document.querySelector(".container");
  contenedor.appendChild(confirmacion);
  
  // Auto-eliminar después de 5 segundos
  setTimeout(() => {
    if (confirmacion.parentNode) {
      confirmacion.remove();
    }
  }, 5000);
}

/* ============================================================
   BOTÓN "ARMAR TOTALIZADOR"
   ============================================================ */

function armarTotalizadorCabal() {
  if (liquidacionesProcesadas.length === 0) {
    alert("Procesá al menos una liquidación primero.");
    return;
  }

  // Crear un libro de trabajo para el totalizador
  const wb = XLSX.utils.book_new();
  const wsData = [];
  
  // Encabezado del totalizador
  wsData.push(["TOTALIZADOR DE LIQUIDACIONES CABAL"]);
  wsData.push([]);
  wsData.push(["N° Liquidación", "Fecha Pago", "Cuenta", "Total Débito", "Total Crédito", "Neto Final"]);
  
  // Sumar los datos de todas las liquidaciones
  let totalDebito = 0;
  let totalCredito = 0;
  let totalNeto = 0;
  
  liquidacionesProcesadas.forEach(liquidacion => {
    const totalDebLiq = liquidacion.tit2.totalVentas.total || 0;
    const totalCreLiq = liquidacion.tit4.totalVentas.total || 0;
    const netoLiq = liquidacion.tit5.totalFinal || 0;
    
    wsData.push([
      liquidacion.encabezado.liquidacionNro || "",
      liquidacion.encabezado.fechaPago || "",
      liquidacion.encabezado.cuenta || "",
      totalDebLiq,
      totalCreLiq,
      netoLiq
    ]);
    
    totalDebito += totalDebLiq;
    totalCredito += totalCreLiq;
    totalNeto += netoLiq;
  });
  
  // Agregar totales
  wsData.push([]);
  wsData.push(["TOTALES", "", "", totalDebito, totalCredito, totalNeto]);
  
  // Crear la hoja
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "Totalizador");
  
  // Descargar el archivo
  const fecha = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Totalizador_CABAL_${fecha}.xlsx`);
}