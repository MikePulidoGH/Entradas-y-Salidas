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
   * 🚀 [ACCIÓN] : Normaliza nombres a "Juan Perez", gestiona fechas en bloque y borrado automático [8-10].
   */
  if (nombreHoja === "SALIDAS") {
    if (filaInicio === 2) {
      const valor = e.range.getValue();
      if (colInicio === 2 && valor === true) { e.range.setValue(false); finalizarJornadaReciente(); }
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
 // --- 4. AJUSTES (LIMPIEZA B-C-I Y FECHADO DINÁMICO) ---
  if (nombreHoja === "AJUSTES") {
    const valorActual = e.range.getValue();

    // A. LIMPIEZA MAESTRA (Celda A1)
    // Regla: Solo limpia columnas B, C e I de las filas 2 a 50
    if (filaInicio === 1 && colInicio === 1 && valorActual === true) {
      hoja.getRange("B2:C50").clearContent(); // Limpia Columnas B y C
      hoja.getRange("I2:I50").clearContent(); // Limpia Columna I
      e.range.setValue(false); // Resetea el botón A1
      SpreadsheetApp.flush();
      return;
    }

    // B. FECHADO DINÁMICO (Columnas B o C - Filas 2 a 50)
    // Si se edita el Código (B) o la Cantidad (C), actualiza o pone la fecha en I
    if (filaInicio >= 2 && filaInicio <= 50 && (colInicio === 2 || colInicio === 3)) {
      const celdaFecha = hoja.getRange(filaInicio, 9); // Columna I
      if (valorActual !== "") {
        celdaFecha.setValue(new Date()).setNumberFormat("dd/mm/yyyy HH:mm");
      } else {
        celdaFecha.clearContent();
      }
    }

    // C. PROCESAR AJUSTE (Checkbox columna A - Filas 2 a 50)
    if (filaInicio >= 2 && filaInicio <= 50 && colInicio === 1 && valorActual === true) {
      const codigo   = hoja.getRange(filaInicio, 2).getValue(); // Columna B
      const valorE   = hoja.getRange(filaInicio, 5).getValue(); // Columna E (Valor solicitado)
      const producto = hoja.getRange(filaInicio, 8).getValue(); // Columna H
      
      if (producto !== "") {
        const ss = SpreadsheetApp.getActive();
        const hojaDestino = ss.getSheetByName("SL AJUSTES");
        
        if (hojaDestino) {
          // Se mantiene el mapeo: Ajuste | Código | Valor E | Producto | Fecha
          hojaDestino.appendRow(["Ajuste", codigo, valorE, producto, new Date()]);
          
          // Aplicar cuadrícula a la fila nueva en destino
          const filaNueva = hojaDestino.getLastRow();
          hojaDestino.getRange(filaNueva, 1, 1, 5)
            .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
          
          ss.toast("✅ Enviado a SL AJUSTES: " + producto);
        }
      }
      e.range.setValue(false); // Apaga el checkbox de origen
    }
    return;
  }
}
/**
 * 🚀 EJECUTA ESTA FUNCIÓN PARA ACTIVAR EL SISTEMA
 */
function instalarTrigger() {
  const ss = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  ScriptApp.newTrigger('onEditInstalado')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
      
  SpreadsheetApp.getUi().alert("✅ Sistema Activado. Ya puedes probar la hoja AJUSTES.");
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

// ==========================================
// 🚀 FUNCIÓN DE ENVÍO CORREGIDA
// ==========================================
function ejecutarEnvioAjuste(producto, cantidad, fila, range, hoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destino = ss.getSheetByName("SL AJUSTES");
  
  if (!destino) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja SL AJUSTES");
    return;
  }

  // Enviamos los datos
  destino.appendRow([
    new Date(), 
    "AJUSTE MANUAL", 
    producto, 
    cantidad, 
    "PROCESADO"
  ]);

  // --- LA CLAVE: LIMPIEZA POST-ENVÍO ---
  hoja.getRange(fila, 2, 1, 2).clearContent(); // Borra B y C
  hoja.getRange(fila, 9).clearContent();      // Borra fecha en I
  range.setValue(false);                      // Desmarca el Check en A
  
  ss.toast("✅ Ajuste de '" + producto + "' enviado y fila limpia.");
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

/**
 * REVISOR ULTRA-RÁPIDO PARA SALIDAS
 * Diseñado para detectar cambios externos al instante.
 */
function ejecutorDeFechasAutomaticas(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("SALIDAS");
  if (!hoja) return;

  // Solo revisamos las últimas 20 filas para que sea "al toque" 
  // y no pierda tiempo procesando miles de filas viejas.
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 3) return;
  
  const filasARevisar = 20; // Ajusta este número según cuántas filas pegues de golpe
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
