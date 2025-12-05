/* ============================================================
   PARSER DEFINITIVO DEL ENCABEZADO CABAL
   ============================================================ */

function parsearEncabezadoCabal(textoOriginal) {

    // 1) Normalizar saltos y espacios
    let texto = textoOriginal
        .replace(/\r/g, "")
        .replace(/\t/g, " ")
        .replace(/ +/g, " ")  // limpia espacios extras
        .trim();

    // 2) Cortar basura desde ENCUADRES FISCALES hasta el primer título real
    texto = limpiarBasuraEncabezado(texto);

    // 3) Separar líneas limpias
    const lineas = texto.split("\n").map(l => l.trim());

    let resultado = {
        fechaPago: "",
        liquidacionNro: "",
        cuenta: ""
    };

    // 4) Buscar cada dato del encabezado (muy estrictamente)
    for (let l of lineas) {

        // FECHA DE PAGO
        if (l.toUpperCase().startsWith("FECHA DE PAGO")) {
            resultado.fechaPago = extraerValor(l);
        }

        // LIQUIDACION NRO
        if (l.toUpperCase().startsWith("LIQUIDACION NRO")) {
            resultado.liquidacionNro = extraerValor(l);
        }

        // CUENTA P/ACREDITAR VENTAS
        if (l.toUpperCase().startsWith("CUENTA P/ACREDITAR VENTAS")) {
            resultado.cuenta = extraerValor(l);
        }
    }

    return resultado;
}

/* ============================================================
   2) ELIMINAR TODA LA BASURA DEL PDF
   ============================================================ */

function limpiarBasuraEncabezado(texto) {

    const lineas = texto.split("\n");

    let limpio = [];
    let modoBasura = false;

    for (let l of lineas) {
        let linea = l.trim();

        // Detectamos el comienzo de BASURA
        if (linea.toUpperCase().startsWith("ENCUADRES FISCALES")) {
            modoBasura = true;
            continue;
        }

        // Detectamos el primer título válido (fin de basura)
        if (
            linea.startsWith("VENTAS CORRESPONDIENTES A CABAL DEBITO") ||
            linea.startsWith("VENTAS CORRESPONDIENTES A TARJETA DE CREDITO") ||
            linea.startsWith("CABAL DEBITO") ||
            linea.startsWith("TARJETA DE CREDITO TOTAL DE VENTAS")
        ) {
            // salimos: ya pasamos el encabezado
            break;
        }

        // Mientras estamos en modo basura → ignoramos líneas
        if (modoBasura) continue;

        // Guardamos líneas válidas del encabezado
        limpio.push(linea);
    }

    return limpio.join("\n");
}

/* ============================================================
   3) EXTRAER VALOR después de ":"
   ============================================================ */

function extraerValor(linea) {
    const partes = linea.split(":");
    if (partes.length < 2) return "";
    return partes[1].trim();
}

/* ============================================================
   TESTER
   ============================================================ */

function testEncabezado() {
    const texto = document.getElementById("textoTest").value;
    const res = parsearEncabezadoCabal(texto);

    document.getElementById("salidaTest").textContent =
        JSON.stringify(res, null, 2);
}
