/* global CustomFunctions */

// ============================================================
// ============= IDENTIDAD FISCAL MX - CUSTOM FUNCTIONS =======
// ============================================================
// Compatible con Excel Desktop (M365) y Excel Online
// Uso: =CURP.GENERAR(...) | =RFC.GENERAR(...) | =CURP.EDAD(...)
// ============================================================

// =================== UTILIDADES =============================

function limpiarTexto(txt) {
  if (!txt) return "";
  const desde = "ÁÉÍÓÚÜÑáéíóúüñ";
  const hacia  = "AEIOUUNAEIOUun".toUpperCase();
  let resultado = txt.toUpperCase().trim();
  for (let i = 0; i < desde.length; i++) {
    resultado = resultado.replaceAll(desde[i], hacia[i]);
  }
  return resultado;
}

function esVocal(letra) {
  return "AEIOU".includes(letra);
}

function primeraVocalInterna(txt) {
  for (let i = 1; i < txt.length; i++) {
    if (esVocal(txt[i])) return txt[i];
  }
  return "X";
}

function consonanteInterna(txt) {
  if (!txt) return "X";
  for (let i = 1; i < txt.length; i++) {
    if (!esVocal(txt[i]) && txt[i] !== "Ñ") return txt[i];
  }
  return "X";
}

function formatFecha(fecha) {
  const d = new Date(fecha);
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yy}${mm}${dd}`;
}

function normalizarSexo(sexo) {
  sexo = limpiarTexto(sexo);
  if (["HOMBRE", "H", "MASCULINO", "M"].includes(sexo)) {
    return sexo === "MUJER" || sexo === "FEMENINO" ? "M" : 
           (sexo === "M" && !["HOMBRE","MASCULINO"].includes(sexo)) ? "M" : "H";
  }
  if (["MUJER", "FEMENINO"].includes(sexo)) return "M";
  if (sexo === "H") return "H";
  if (sexo === "M") return "M";
  return "";
}

function normalizarEstadoCURP(estado) {
  estado = limpiarTexto(estado);
  const dic = {
    "AGUASCALIENTES": "AS", "BAJA CALIFORNIA": "BC", "BAJA CALIFORNIA SUR": "BS",
    "CAMPECHE": "CC", "COAHUILA": "CL", "COLIMA": "CM", "CHIAPAS": "CS",
    "CHIHUAHUA": "CH", "CIUDAD DE MEXICO": "DF", "CDMX": "DF",
    "DISTRITO FEDERAL": "DF", "DURANGO": "DG", "GUANAJUATO": "GT",
    "GUERRERO": "GR", "HIDALGO": "HG", "JALISCO": "JC", "MEXICO": "MC",
    "ESTADO DE MEXICO": "MC", "EDO MEX": "MC", "MICHOACAN": "MN",
    "MORELOS": "MS", "NAYARIT": "NT", "NUEVO LEON": "NL", "OAXACA": "OC",
    "PUEBLA": "PL", "QUERETARO": "QT", "QUINTANA ROO": "QR",
    "SAN LUIS POTOSI": "SP", "SINALOA": "SL", "SONORA": "SR",
    "TABASCO": "TC", "TAMAULIPAS": "TS", "TLAXCALA": "TL",
    "VERACRUZ": "VZ", "YUCATAN": "YN", "ZACATECAS": "ZS",
    "EXTRANJERO": "NE", "NACIDO EN EL EXTRANJERO": "NE"
  };
  return dic[estado] || estado;
}

// =================== CURP ===================================

function curpValorCaracter(ch) {
  const tabla = "0123456789ABCDEFGHIJKLMNÑOPQRSTUVWXYZ";
  return tabla.indexOf(ch);
}

/**
 * Genera la CURP de una persona.
 * @customfunction
 * @param {string} nombre Nombre(s) de la persona
 * @param {string} paterno Apellido paterno
 * @param {string} materno Apellido materno (dejar vacío si no tiene)
 * @param {string} fechaNacimiento Fecha de nacimiento (ej: "1990-05-15")
 * @param {string} sexo "H", "Hombre", "M" o "Mujer"
 * @param {string} estado Estado de nacimiento (ej: "Jalisco")
 * @returns {string} CURP de 18 caracteres
 */
function GENERAR_CURP(nombre, paterno, materno, fechaNacimiento, sexo, estado) {
  try {
    nombre  = limpiarTexto(nombre);
    paterno = limpiarTexto(paterno);
    materno = limpiarTexto(materno || "");
    sexo    = normalizarSexo(sexo);
    estado  = normalizarEstadoCURP(estado);

    if (!paterno) return "ERROR: Falta apellido paterno";
    if (!sexo)   return "ERROR: Sexo inválido (usa H/M/Hombre/Mujer)";
    if (!estado) return "ERROR: Estado inválido";

    const fecha = new Date(fechaNacimiento);
    if (isNaN(fecha)) return "ERROR: Fecha inválida";

    const yy = String(fecha.getFullYear()).slice(-2);
    const mm = String(fecha.getMonth() + 1).padStart(2, "0");
    const dd = String(fecha.getDate()).padStart(2, "0");

    let base =
      paterno[0] +
      primeraVocalInterna(paterno) +
      (materno ? materno[0] : "X") +
      nombre[0] +
      yy + mm + dd +
      sexo +
      estado +
      consonanteInterna(paterno) +
      consonanteInterna(materno || "") +
      consonanteInterna(nombre);

    // Diferenciador de siglo
    const diferenciador = fecha.getFullYear() < 2000 ? "0" : "A";
    base += diferenciador;

    // Dígito verificador
    let suma = 0;
    for (let i = 0; i < base.length; i++) {
      suma += curpValorCaracter(base[i]) * (19 - i);
    }
    const digito = (10 - (suma % 10)) % 10;

    return base + digito;
  } catch (e) {
    return "ERROR: " + e.message;
  }
}

// =================== RFC ====================================

function rfcFiltrarNombres(nombre, paterno, materno) {
  const particulas = ["DE ", "DEL ", "LA ", "LOS ", "LAS ", "Y ", "MC ", "MAC ", "VON ", "VAN "];
  const prefijosNombre = ["JOSE ", "MARIA ", "J ", "MA "];

  for (const p of particulas) {
    nombre  = nombre.replaceAll(p, "");
    paterno = paterno.replaceAll(p, "");
    materno = materno.replaceAll(p, "");
  }
  nombre  = nombre.replaceAll(".", "").replaceAll(",", "").trim();
  paterno = paterno.replaceAll(".", "").replaceAll(",", "").trim();
  materno = materno.replaceAll(".", "").replaceAll(",", "").trim();

  if (nombre.includes(" ")) {
    for (const p of prefijosNombre) {
      nombre = nombre.replaceAll(p, "");
    }
  }

  return { nombre: nombre.trim(), paterno: paterno.trim(), materno: materno.trim() };
}

function rfcQuitarProhibidas(rfc) {
  const prohibidas = [
    "BUEI","BUEY","CACA","CACO","CAGA","CAGO","CAKA","CAKO","COGE","COJA",
    "KOGE","KOJO","KAKA","KULO","MAME","MAMO","MEAR","MEAS","MEON","MION",
    "COJE","COJI","COJO","CULO","FETO","GUEY","JOTO","KACA","KACO","KAGA",
    "KAGO","MOCO","MULA","PEDA","PEDO","PENE","PUTA","PUTO","QULO","RATA","RUIN"
  ];
  if (prohibidas.includes(rfc.slice(0, 4))) {
    return rfc.slice(0, 3) + "X" + rfc.slice(4);
  }
  return rfc;
}

function rfcHomoclave(nombre, paterno, materno) {
  const nombreComp = `${paterno} ${materno} ${nombre}`;
  const charsHc = "123456789ABCDEFGHIJKLMNPQRSTUVWXYZ";
  let cadena = "0";

  for (const ch of nombreComp) {
    const code = ch.charCodeAt(0);
    if (ch === " " || ch === "-") { cadena += "00"; }
    else if (ch === "Ñ" || ch === "Ü") { cadena += "10"; }
    else if (code >= 65 && code <= 73) { cadena += String(code - 54); }  // A-I
    else if (code >= 74 && code <= 82) { cadena += String(code - 53); }  // J-R
    else if (code >= 83 && code <= 90) { cadena += String(code - 51); }  // S-Z
    else if (ch >= "0" && ch <= "9")   { cadena += ch.padStart(2, "0"); }
  }

  let suma = 0;
  for (let i = 0; i < cadena.length - 1; i++) {
    const n1 = parseInt(cadena.slice(i, i + 2));
    const n2 = parseInt(cadena[i + 1]);
    suma += n1 * n2;
  }

  const tres = parseInt(String(suma).slice(-3)) || 0;
  const quo  = Math.floor(tres / 34);
  const rem  = tres % 34;

  return charsHc[quo] + charsHc[rem];
}

function rfcDigitoVerificador(rfc) {
  const chars = "0123456789ABCDEFGHIJKLMN&OPQRSTUVWXYZ ";
  let suma = 0;
  for (let i = 0; i < rfc.length; i++) {
    const ch  = rfc[i] === " " ? "*" : rfc[i];
    const idx = chars.indexOf(ch);
    if (idx < 0) return "0";
    suma += idx * (14 - i);
  }
  const mod = suma % 11;
  if (mod === 0) return "0";
  const dv = 11 - mod;
  return dv > 9 ? "A" : String(dv);
}

/**
 * Genera el RFC de una persona física.
 * @customfunction
 * @param {string} nombre Nombre(s) de la persona
 * @param {string} paterno Apellido paterno
 * @param {string} materno Apellido materno
 * @param {string} fechaNacimiento Fecha de nacimiento (ej: "1990-05-15")
 * @returns {string} RFC de 13 caracteres
 */
function GENERAR_RFC(nombre, paterno, materno, fechaNacimiento) {
  try {
    nombre  = limpiarTexto(nombre);
    paterno = limpiarTexto(paterno);
    materno = limpiarTexto(materno || "");

    const fecha = new Date(fechaNacimiento);
    if (isNaN(fecha)) return "ERROR: Fecha inválida";

    const yy = String(fecha.getFullYear()).slice(-2);
    const mm = String(fecha.getMonth() + 1).padStart(2, "0");
    const dd = String(fecha.getDate()).padStart(2, "0");
    const strFecha = yy + mm + dd;

    // Guardar originales para homoclave
    const nomOrig = nombre;
    const patOrig = paterno;
    const matOrig = materno;

    const filtrado = rfcFiltrarNombres(nombre, paterno, materno);
    nombre  = filtrado.nombre;
    paterno = filtrado.paterno;
    materno = filtrado.materno;

    let rfc;
    if (paterno && materno) {
      rfc = paterno[0] + primeraVocalInterna(paterno) +
            (materno[0] || "X") + nombre[0] + strFecha;
    } else {
      rfc = (paterno + materno + nombre).slice(0, 4) + strFecha;
    }

    rfc = rfcQuitarProhibidas(rfc);
    rfc = rfc + rfcHomoclave(nomOrig, patOrig, matOrig);

    return rfc + rfcDigitoVerificador(rfc);
  } catch (e) {
    return "ERROR: " + e.message;
  }
}

// =================== EXTRAS (leer CURP) =====================

/**
 * Extrae la fecha de nacimiento desde una CURP.
 * @customfunction
 * @param {string} curp CURP de 18 caracteres
 * @returns {string} Fecha en formato AAAA-MM-DD
 */
function CURP_FECHA(curp) {
  curp = limpiarTexto(curp);
  if (curp.length !== 18) return "ERROR: CURP inválida";
  const anio = parseInt(curp.slice(4, 6));
  const mes  = curp.slice(6, 8);
  const dia  = curp.slice(8, 10);
  const siglo = /[0-9]/.test(curp[16]) ? 1900 : 2000;
  return `${siglo + anio}-${mes}-${dia}`;
}

/**
 * Extrae el sexo desde una CURP.
 * @customfunction
 * @param {string} curp CURP de 18 caracteres
 * @returns {string} "Hombre" o "Mujer"
 */
function CURP_SEXO(curp) {
  curp = limpiarTexto(curp);
  if (curp.length !== 18) return "ERROR: CURP inválida";
  return curp[10] === "H" ? "Hombre" : curp[10] === "M" ? "Mujer" : "ERROR";
}

/**
 * Calcula la edad actual desde una CURP.
 * @customfunction
 * @param {string} curp CURP de 18 caracteres
 * @returns {number} Edad en años
 */
function CURP_EDAD(curp) {
  const fechaStr = CURP_FECHA(curp);
  if (fechaStr.startsWith("ERROR")) return fechaStr;
  const nac  = new Date(fechaStr);
  const hoy  = new Date();
  let edad = hoy.getFullYear() - nac.getFullYear();
  if (hoy < new Date(hoy.getFullYear(), nac.getMonth(), nac.getDate())) edad--;
  return edad;
}

/**
 * Extrae el estado de nacimiento desde una CURP.
 * @customfunction
 * @param {string} curp CURP de 18 caracteres
 * @returns {string} Nombre del estado
 */
function CURP_ESTADO(curp) {
  curp = limpiarTexto(curp);
  if (curp.length !== 18) return "ERROR: CURP inválida";
  const cod = curp.slice(11, 13);
  const estados = {
    AS:"Aguascalientes", BC:"Baja California", BS:"Baja California Sur",
    CC:"Campeche", CS:"Chiapas", CH:"Chihuahua", CL:"Coahuila", CM:"Colima",
    DF:"Ciudad de México", DG:"Durango", GT:"Guanajuato", GR:"Guerrero",
    HG:"Hidalgo", JC:"Jalisco", MC:"México", MN:"Michoacán", MS:"Morelos",
    NT:"Nayarit", NL:"Nuevo León", OC:"Oaxaca", PL:"Puebla", QT:"Querétaro",
    QR:"Quintana Roo", SP:"San Luis Potosí", SL:"Sinaloa", SR:"Sonora",
    TC:"Tabasco", TS:"Tamaulipas", TL:"Tlaxcala", VZ:"Veracruz",
    YN:"Yucatán", ZS:"Zacatecas", NE:"Extranjero"
  };
  return estados[cod] || "Desconocido";
}

// =================== REGISTRO DE FUNCIONES ==================

CustomFunctions.associate("CURP.GENERAR", GENERAR_CURP);
CustomFunctions.associate("RFC.GENERAR",  GENERAR_RFC);
CustomFunctions.associate("CURP.FECHA",   CURP_FECHA);
CustomFunctions.associate("CURP.SEXO",    CURP_SEXO);
CustomFunctions.associate("CURP.EDAD",    CURP_EDAD);
CustomFunctions.associate("CURP.ESTADO",  CURP_ESTADO);
