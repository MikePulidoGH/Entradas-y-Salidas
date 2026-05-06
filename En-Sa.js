/**
	 * SISTEMA SALIDAS v11.0 - VERSIÓN FINAL ESTABILIZADA
 */

// ── WEBHOOK ÚNICO ──────────────────────────────────────────────────────────
const WEBHOOK_DISCORD = "https://discordapp.com/api/webhooks/1463548831915442322/Fd4MEYgwVtxz0vo8o4AR11CeMP_NsL3CTyXp83ai9Emv4vp30So4STq14a45PNkvvUiN";

// ── ON EDIT ────────────────────────────────────────────────────────────────
/**
 * 🧠 [ACTIVADOR] : onEditInstalado(e)
 * 📍 [HOJA]      : TODAS (Pedidos, Salidas, Entradas, Ajustes)
 * 🚀 [ACCIÓN]    : El "semáforo" principal que detecta cambios y activa la lógica correcta [1, 3].
 */
function onEditInstalado(e) {
  if (!e || !e.range) return;
												
  const hoja       = e.range.getSheet();
  const nombreHoja = hoja.getName();
  const filaInicio = e.range.getRow();
  const colInicio  = e.range.getColumn();
  const numFilas   = e.range.getNumRows();
  const numCols    = e.range.getNumColumns();

  // --- 1. PEDIDOS ---
  /**
   * 📍 [HOJA]   : "PEDIDOS"
   * 🎮 [BOTÓN]  : Checkbox en Columna A (Fila 1) [4, 5].
   * 🚀 [ACCIÓN] : Dispara la transferencia masiva de pedidos hacia la hoja ENTRADAS [4, 5].
   */
  if (nombreHoja === "PEDIDOS" && colInicio === 1 && numFilas === 1) {
    if (e.range.getValue() === true) {
      try {
        enviarPedidosAEntradas();
      } catch (error) {
        SpreadsheetApp.getUi().alert("❌ Error al procesar Pedidos: " + error.message);
        e.range.setValue(false);
      }
    }
    return;
  }

  // --- 2. SALIDAS (NORMALIZACIÓN Y BLOQUES) ---
  /**
   * 📍 [HOJA]   : "SALIDAS"
   * 🎮 [BOTONES]: Checkboxes en Fila 2 (Cierre de día, Extender tabla, Cuadrícula) [6, 7].
   */
  if (nombreHoja === "SALIDAS") {
    if (filaInicio === 2) {
      const valor = e.range.getValue();
      if (colInicio === 2 && valor === true) { e.range.setValue(false); finalizarJornadaReciente(); }
/* 🚀 [ACCIÓN] : Normaliza nombres a "Juan Perez", gestiona fechas en bloque y borrado automático [8-10].*/
      if (colInicio === 6 && valor === true) { e.range.setValue(false); extenderTabla150(hoja); }
      if (colInicio === 7 && valor === true) { e.range.setValue(false); aplicarCuadriculaTotal(hoja); }
      return;
    }

    if (filaInicio > 2 && colInicio <= 2) {
      const rangoEfectivo = hoja.getRange(filaInicio, 1, numFilas, 5); 
      const datosRango = rangoEfectivo.getValues();
      const nuevasFechas = [];
      const nuevosValoresA = [];
      let huboCambios = false;

      for (let i = 0; i < datosRango.length; i++) {
        let valA = datosRango[i][0] ? datosRango[i][0].toString() : "";
        const valB = datosRango[i][1]; 
        const fechaActual = datosRango[i][4]; 
        let resultadoFecha = fechaActual;

        // REGLA SALIDAS: Iniciales Mayúscula (Ej: juan perez -> Juan Perez)
        if (valA !== "") {
          valA = valA.toLowerCase().split(' ')
            .map(palabra => palabra.charAt(0).toUpperCase() + palabra.slice(1))
            .join(' ').trim();
        }
        nuevosValoresA.push([valA]);

        if ((valA !== "" || valB !== "") && (!fechaActual || fechaActual === "")) {
          resultadoFecha = new Date();
          huboCambios = true;
        } else if (valA === "" && valB === "" && fechaActual !== "") {
          resultadoFecha = "";
          huboCambios = true;
        }
        nuevasFechas.push([resultadoFecha]);
      }

      // Aplicar transformación en columna A
      hoja.getRange(filaInicio, 1, numFilas, 1).setValues(nuevosValoresA);

      if (huboCambios) {
        hoja.getRange(filaInicio, 5, numFilas, 1)
            .setValues(nuevasFechas)
            .setNumberFormat("dd/mm/yyyy HH:mm");
        aplicarFormatoFila(hoja.getRange(filaInicio, 1, numFilas, 7));
      }
    }
    return;
  }
// --- 3. ENTRADAS (NORMALIZACIÓN Y DISCORD) ---
if (nombreHoja === "ENTRADAS") {
  // Definimos el rango de acción: Columnas A (Proveedor) o B (Código)
  if (colInicio <= 2) {
    const numFilasEditadas = numFilas;
    const rangoA = hoja.getRange(filaInicio, 1, numFilasEditadas, 1);
    const rangoB = hoja.getRange(filaInicio, 2, numFilasEditadas, 1);
    const rangoE = hoja.getRange(filaInicio, 5, numFilasEditadas, 1);
    
    const valoresA = rangoA.getValues();
    const valoresB = rangoB.getValues();
    const valoresE = rangoE.getValues();

    const nuevosValoresA = [];
    const nuevasFechas = [];
    const lineasDiscord = [];

    for (let i = 0; i < numFilasEditadas; i++) {
      let nombre = valoresA[i][0] ? valoresA[i][0].toString().toUpperCase().trim() : "";
      let codigo = valoresB[i][0] ? valoresB[i][0].toString().trim() : "";
      let fechaExistente = valoresE[i][0];

      nuevosValoresA.push([nombre]);

      // LÓGICA DE FECHADO: 
      // Si hay CÓDIGO y NO hay fecha -> Ponemos fecha nueva y notificamos
      if (codigo !== "" && (!fechaExistente || fechaExistente === "")) {
        const ahora = new Date();
        nuevasFechas.push([ahora]);
        
        // Obtener cantidad de la columna C (3)
        const cantidad = hoja.getRange(filaInicio + i, 3).getValue() || "0";
        lineasDiscord.push(`📦 **${codigo.toUpperCase()}**\n> Proveedor: **${nombre || "—"}**\n> Cantidad: **${cantidad}**`);
      } 
      // Si el CÓDIGO se borró -> Limpiamos la fecha
      else if (codigo === "" && fechaExistente !== "") {
        nuevasFechas.push([""]);
      }
      // De lo contrario, mantenemos lo que hay
      else {
        nuevasFechas.push([fechaExistente]);
      }
    }

    // Aplicar cambios en bloque para optimizar velocidad
    rangoA.setValues(nuevosValoresA);
    rangoE.setValues(nuevasFechas).setNumberFormat("dd/mm/yyyy HH:mm");
  }
  return;
}

  /**
   * 📍 [HOJA]   : "AJUSTES"
   * 🚀 [ACCIÓN] : Ejecuta Limpieza Maestra desde A1 y fechado dinámico en Columna I
   */
  // --- 4. AJUSTES (LIMPIEZA Y TRANSFERENCIA) ---
 if (nombreHoja === "AJUSTES") {
  
  // A. LIMPIEZA MAESTRA (Celda A1)
  if (filaInicio === 1 && colInicio === 1 && e.range.getValue() === true) {
    hoja.getRange("B2:C50").clearContent();
    hoja.getRange("I2:I50").clearContent(); // <-- AGREGAR ESTA LÍNEA [1, 3]
    e.range.setValue(false);
    SpreadsheetApp.flush();
    return; // Finaliza tras limpiar
  }

  // B. FECHADO DINÁMICO (Columnas B o C)
  if (filaInicio >= 2 && (colInicio === 2 || colInicio === 3)) {
    const valorB = hoja.getRange(filaInicio, 2).getValue();
    const valorC = hoja.getRange(filaInicio, 3).getValue();
    const celdaFecha = hoja.getRange(filaInicio, 9); // Columna I [4, 8]
    
    if (valorB !== "" || valorC !== "") {
      celdaFecha.setValue(new Date()).setNumberFormat("dd/mm/yyyy HH:mm");
    } else {
      celdaFecha.clearContent();
    }
  }

  // C. ENVÍO A "SL AJUSTES" (Checkboxes desde Fila 3, Columna A)
  if (filaInicio >= 2 && colInicio === 1) {
    const valorCheckbox = e.range.getValue();
    if (valorCheckbox === true) {
      // Esta función es la que mueve los datos a la otra hoja [6, 9]
      procesarHojaAjustes(hoja, filaInicio, colInicio, valorCheckbox, e.range);
    }
  }
  return; 
}
}
// ── ON OPEN ────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ GESTIÓN')
    .addItem('Cerrar y Agrupar Día Reciente', 'finalizarJornadaReciente')
    .addToUi();
}

// ── PEDIDOS → ENTRADAS + DISCORD ──────────────────────────────────────────


// ── AUXILIAR DISCORD ──────────────────────────────────────────────────────
/**
 * Envía un embed a Discord.
 * @param {string} titulo  - Título del embed
 * @param {string} detalle - Cuerpo del mensaje (acepta markdown de Discord)
 * @param {number} color   - Color del embed (decimal). Verde=3066993, Rojo=15158332
 */
function enviarADiscord(titulo, detalle, color) {
  const payload = JSON.stringify({
    "embeds": [{
      "title": titulo,
      "description": detalle,
      "color": color,
      "footer": { "text": "Gestión de Stock | " + new Date().toLocaleString("es-AR") }
    }]
  });

  try {
    UrlFetchApp.fetch(WEBHOOK_DISCORD.trim(), {
      "method": "post",
      "contentType": "application/json",
      "payload": payload,
      "muteHttpExceptions": true
    });
  } catch (e) {
    console.log("Error Discord: " + e.toString());
  }
}

// ── SALIDAS: FUNCIONES DE SOPORTE ─────────────────────────────────────────
function finalizarJornadaReciente() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const hoja       = ss.getSheetByName("SALIDAS");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 2) return;

  const datosB  = hoja.getRange(3, 2, ultimaFila - 2).getValues();
  let filaFinal = -1;
  let fechaE    = new Date();

  for (let i = datosB.length - 1; i >= 0; i--) {
    if (datosB[i][0] !== "") {
      filaFinal = i + 3;
      const valFecha = hoja.getRange(filaFinal, 5).getValue();
      if (valFecha instanceof Date) fechaE = valFecha;
      break;
    }
  }
  if (filaFinal === -1) return;

  let filaInicio = filaFinal;
  for (let j = filaFinal; j >= 3; j--) {
    if (hoja.getRange(j, 2).getValue() !== "") filaInicio = j;
    else break;
  }

  try {
    hoja.getRange(filaInicio, 1, (filaFinal - filaInicio) + 1).shiftRowGroupDepth(1);

    hoja.insertRowAfter(filaFinal);
    const filaGris  = filaFinal + 1;
    const rangoGris = hoja.getRange(filaGris, 1, 1, 7);
    try { rangoGris.shiftRowGroupDepth(-1); } catch(err){}
    hoja.getRange(filaGris, 5).setValue(fechaE).setNumberFormat("dd/mm/yyyy").setFontWeight("bold");
    rangoGris.setBackground("#f3f3f3");
    aplicarFormatoFila(rangoGris);

    hoja.insertRowsAfter(filaGris, 5);
    const rangoAire = hoja.getRange(filaGris + 1, 1, 5, 7);
    hoja.getRange(3, 1, 1, 7).copyTo(rangoAire);
    rangoAire.clearContent();
    aplicarFormatoFila(rangoAire);

    hoja.collapseAllRowGroups();
    ss.toast("Cierre completado. Formato listo.", "✅");
  } catch (err) {}
}

function extenderTabla150(hoja) {
  const max       = hoja.getMaxRows();
  const filaMolde = hoja.getRange(3, 1, 1, 7); // Toma la fila 3 como molde (incluye tus fórmulas)
  
  hoja.insertRowsAfter(max, 150);
  
  const destino = hoja.getRange(max + 1, 1, 150, 7);
  
  // 1. Copiamos el molde (esto arrastra las fórmulas de la fila 3 a las 150 nuevas)
  filaMolde.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  // 2. Limpiamos solo las columnas de DATOS (A, B, C) -> Columnas 1 a 3
  destino.offset(0, 0, 150, 3).clearContent();
  
  // 3. Limpiamos la columna de FECHA (E) -> Columna 5
  // offset(filas, columnas, numFilas, numColumnas)
  // El 4 indica que se mueve 4 columnas a la derecha (llega a la E)
  destino.offset(0, 4, 150, 1).clearContent();
  
  // NOTA: La columna D (índice 3 del offset) NO se toca, por lo tanto la fórmula persiste.
  
  aplicarFormatoFila(destino);
}

function aplicarFormatoFila(rango) {
  rango.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
}

function aplicarCuadriculaTotal(hoja) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 3) return;
  const rangoTotal = hoja.getRange(3, 1, ultimaFila - 2, 7);
  rangoTotal.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  const rangoFormulasD = hoja.getRange(3, 4, ultimaFila - 2, 1);
  hoja.getRange(3, 4).copyTo(rangoFormulasD);
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Cuadrícula y fórmulas restauradas hasta la columna G", "SISTEMA");
}

function procesarHojaAjustes(hoja, fila, col, valor, celda) {
  if (col === 1 && valor === true) {
    const codigo  = hoja.getRange(fila, 2).getValue();
    const cant    = hoja.getRange(fila, 5).getValue();
    const prod    = hoja.getRange(fila, 8).getValue();
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const destino = ss.getSheetByName("SL AJUSTES");
    if (destino && codigo && cant) {
      destino.appendRow(["Ajuste", codigo, cant, prod, new Date()]);
      aplicarFormatoFila(destino.getRange(destino.getLastRow(), 1, 1, 5));
      celda.setValue(false);
      ss.toast("Enviado a AJUSTES", "✅");
    }
  }
}

// ── UTILIDADES ────────────────────────────────────────────────────────────
/**
 * BUSCAR ÚLTIMA FILA REAL
 * Encuentra la última fila que contiene un producto real, 
 * ignorando las filas que solo tienen la fecha de cierre en la columna E.
 */
function buscarUltimaFilaReal(hoja) {
  // Traemos las columnas A (Código) y E (Fecha/Cierre)
  const rango = hoja.getRange("A:E").getValues();
  
  // Recorremos de abajo hacia arriba
  for (let i = rango.length - 1; i >= 0; i--) {
    const colA = String(rango[i][0]).trim();
    const colE = rango[i][4];
    
    // Si la fila tiene algo en A (Producto) o algo en E (Cierre/Fecha)
    // significa que esa fila está OCUPADA.
    if (colA !== "" || (colE !== "" && colE !== null)) {
      return i + 1; // Retorna esa fila como la última ocupada
    }
  }
  return 2; // Si está vacía, devuelve fila 2 para empezar en la 3
}


function finalizarJornadaReciente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("SALIDAS");
  if (!hoja) return;

  const FILA_INICIO = 3;
  const COL_B = 2; // Columna B: referencia de fila con movimiento
  const COL_D = 4; // Columna D: marca DIA CERRADO
  const COL_E = 5; // Columna E: fecha
  const TOTAL_COLS = 7;
  const MARCA_CIERRE = "DIA CERRADO";

  let ultimaFila = hoja.getLastRow();
  if (ultimaFila < FILA_INICIO) return;

  // Lee TODO de una vez para evitar lecturas inconsistentes
  const cantidad = ultimaFila - FILA_INICIO + 1;
  const dataB = hoja.getRange(FILA_INICIO, COL_B, cantidad, 1).getValues();           // valores reales
  const dataEVisible = hoja.getRange(FILA_INICIO, COL_E, cantidad, 1).getDisplayValues(); // fecha visible
  const dataEReal = hoja.getRange(FILA_INICIO, COL_E, cantidad, 1).getValues();        // fecha real

  // Normaliza fecha visible a "dd/mm/yyyy"
  const fechaVisible = (txt) => {
    const s = String(txt || "").trim();
    if (!s) return "";
    const soloFecha = s.split(" ")[0].trim(); // quita hora si existe
    return soloFecha;
  };

  // Armamos bloques por fecha visible y B no vacía
  const bloques = [];
  let i = 0;
  while (i < cantidad) {
    const b = dataB[i][0];
    const f = fechaVisible(dataEVisible[i][0]);

    if (b === "" || f === "") {
      i++;
      continue;
    }

    const inicio = i;
    let fin = i;

    while (fin + 1 < cantidad) {
      const bSig = dataB[fin + 1][0];
      const fSig = fechaVisible(dataEVisible[fin + 1][0]);
      if (bSig !== "" && fSig === f) fin++;
      else break;
    }

    bloques.push({ inicioAbs: FILA_INICIO + inicio, finAbs: FILA_INICIO + fin, fechaKey: f });
    i = fin + 1;
  }

  if (bloques.length === 0) {
    ss.toast("No hay bloques para cerrar.", "ℹ️");
    return;
  }

  // Procesar de abajo hacia arriba para que inserciones no desplacen pendientes
  let cierres = 0;
  for (let k = bloques.length - 1; k >= 0; k--) {
    const { inicioAbs, finAbs } = bloques[k];
    const filaSiguiente = finAbs + 1;

    // Si ya está cerrada, saltar
    const marca = String(hoja.getRange(filaSiguiente, COL_D).getValue()).trim().toUpperCase();
    if (marca === MARCA_CIERRE) continue;

    // Fecha para fila gris: usar fecha real del fin de bloque; si no sirve, hoy
    let fechaE = hoja.getRange(finAbs, COL_E).getValue();
    if (!(fechaE instanceof Date) || isNaN(fechaE)) fechaE = new Date();

    // Agrupar bloque
    hoja.getRange(inicioAbs, 1, (finAbs - inicioAbs) + 1).shiftRowGroupDepth(1);

    // Insertar fila gris
    hoja.insertRowAfter(finAbs);
    const filaGris = finAbs + 1;
    const rangoGris = hoja.getRange(filaGris, 1, 1, TOTAL_COLS);

    try { rangoGris.shiftRowGroupDepth(-1); } catch (err) {}

    hoja.getRange(filaGris, COL_D).setValue(MARCA_CIERRE).setFontWeight("bold");
    hoja.getRange(filaGris, COL_E).setValue(fechaE).setNumberFormat("dd/mm/yyyy").setFontWeight("bold");
    rangoGris.setBackground("#f3f3f3");
    aplicarFormatoFila(rangoGris);

    // Insertar 5 filas "aire"
    hoja.insertRowsAfter(filaGris, 5);
    const rangoAire = hoja.getRange(filaGris + 1, 1, 5, TOTAL_COLS);
    hoja.getRange(FILA_INICIO, 1, 1, TOTAL_COLS).copyTo(rangoAire);
    rangoAire.clearContent();
    aplicarFormatoFila(rangoAire);

    cierres++;
  }

  hoja.collapseAllRowGroups();
  ss.toast(`Cierre completado. Jornadas cerradas: ${cierres}`, "✅");
}


function corregirFechasEntradas() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("ENTRADAS");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const rango   = hoja.getRange(2, 2, ultimaFila - 1, 4);
  const valores = rango.getValues();
  const fechaHoy = new Date();
  let cambios = false;

  for (let i = 0; i < valores.length; i++) {
    const codigo = valores[i][0];
    const fecha  = valores[i][3];
    if (codigo === "" && fecha !== "") { valores[i][3] = ""; cambios = true; }
    else if (codigo !== "" && fecha === "") { valores[i][3] = fechaHoy; cambios = true; }
  }

  if (cambios) {
    rango.setValues(valores);
    SpreadsheetApp.getUi().alert("✅ Fechas sincronizadas.");
  }
}

function limpiadorAutomaticoDeFechas() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("ENTRADAS");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const rango   = hoja.getRange(2, 2, ultimaFila - 1, 4);
  const valores = rango.getValues();

  for (let i = 0; i < valores.length; i++) {
    const codigo = valores[i][0].toString().trim();
    const fecha  = valores[i][3];
    if (codigo === "" && fecha !== "") {
      hoja.getRange(i + 2, 5).clearContent();
    }
  }
}

function testFinalDiscord() {
  const payload = JSON.stringify({ "content": "📢 TEST: Si recibes esto, el Webhook es compatible." });
  try {
    const res = UrlFetchApp.fetch(WEBHOOK_DISCORD.trim(), {
      "method": "post",
      "contentType": "application/json",
      "payload": payload,
      "muteHttpExceptions": true
    });
    console.log("Código: " + res.getResponseCode());
  } catch (e) {
    console.error("Error: " + e.toString());
  }
}
  // Borra triggers viejos de onEdit para no duplicar
function instalarTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let t of triggers) {
    if (t.getHandlerFunction() === "onEditInstalado") {
      ScriptApp.deleteTrigger(t);
    }
  }
  // Instala el nuevo
  ScriptApp.newTrigger("onEditInstalado")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}
/**
 * REVISOR ULTRA-RÁPIDO PARA SALIDAS
 * Diseñado para detectar cambios externos al instante.
 */
function ejecutorDeFechasAutomaticas(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("SALIDAS");
  if (!hoja) return;

  // Solo revisamos las últimas 350 filas para que sea "al toque" 
  // y no pierda tiempo procesando miles de filas viejas.
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 3) return;
  
  const filasARevisar = 350; // Ajusta este número según cuántas filas pegues de golpe
  const inicio = Math.max(3, ultimaFila - filasARevisar + 1);
  const cantidad = ultimaFila - inicio + 1;

  const rango = hoja.getRange(inicio, 1, cantidad, 5);
  const valores = rango.getValues();
  const fechasNuevas = [];
  let hayCambios = false;

  for (let i = 0; i < valores.length; i++) {
    const colA = valores[i][0];
    const colB = valores[i][1];
    const fechaExistente = valores[i][4];

    if ((colA !== "" || colB !== "") && (!fechaExistente || fechaExistente === "")) {
      fechasNuevas.push([new Date()]);
      hayCambios = true;
    } else {
      fechasNuevas.push([fechaExistente]);
    }
  }

  if (hayCambios) {
    hoja.getRange(inicio, 5, fechasNuevas.length, 1)
        .setValues(fechasNuevas)
        .setNumberFormat("dd/mm/yyyy HH:mm");
    
    // Forzamos a la hoja a mostrar los cambios inmediatamente
    SpreadsheetApp.flush(); 
  }
}
/**
 * REVISOR DE COLUMNA A -> FECHA EN E
 * Esta función busca filas con datos en A que no tengan fecha en E y la asigna.
 */
function revisorAutomaticoColumnaA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("SALIDAS");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 3) return; // Empezamos en fila 3 según tu estructura [cite: 38]

  // Leemos Columnas A hasta E (1 a 5)
  const rango = hoja.getRange(3, 1, ultimaFila - 2, 5);
  const datos = rango.getValues();
  const fechasNuevas = [];
  let huboCambios = false;

  for (let i = 0; i < datos.length; i++) {
    const valorA = datos[i][0];       // Columna A
    const fechaActual = datos[i][4];  // Columna E
    
    // Si A tiene datos Y la fecha en E está vacía
    if (valorA !== "" && (!fechaActual || fechaActual === "")) {
      fechasNuevas.push([new Date()]);
      huboCambios = true;
    } else {
      // Mantenemos lo que ya estaba (ya sea la fecha o el vacío)
      fechasNuevas.push([fechaActual]);
    }
  }

  if (huboCambios) {
    // Escribimos solo en la columna E (columna 5) de una sola vez
    hoja.getRange(3, 5, fechasNuevas.length, 1)
        .setValues(fechasNuevas)
        .setNumberFormat("dd/mm/yyyy HH:mm");
    
    console.log("✅ Fechas sincronizadas correctamente.");
  }
}
