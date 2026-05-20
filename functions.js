/* global CustomFunctions */

function limpiarTexto(txt) {
  if (!txt) return "";
  const mapa = {
    "Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ü":"U","Ñ":"N",
    "á":"A","é":"E","í":"I","ó":"O","ú":"U","ü":"U","ñ":"N"
  };
  let r = txt.toUpperCase().trim();
  for (const [k,v] of Object.entries(mapa)) r = r.split(k).join(v);
  return r;
}

function esVocal(l) { return "AEIOU".includes(l); }

function primeraVocalInterna(txt) {
  for (let i = 1; i < txt.length; i++) if (esVocal(txt[i])) return txt[i];
  return "X";
}

function consonanteInterna(txt) {
  if (!txt) return "X";
  for (let i = 1; i < txt.length; i++)
    if (!esVocal(txt[i]) && txt[i] !== "N") return txt[i];
  return "X";
}

function normalizarSexo(s) {
  s = limpiarTexto(s);
  if (["H","HOMBRE","MASCULINO"].includes(s)) return "H";
  if (["M","MUJER","FEMENINO"].includes(s))   return "M";
  return "";
}

function normalizarEstado(e) {
  e = limpiarTexto(e);
  const d = {
    "AGUASCALIENTES":"AS","BAJA CALIFORNIA":"BC","BAJA CALIFORNIA SUR":"BS",
    "CAMPECHE":"CC","COAHUILA":"CL","COLIMA":"CM","CHIAPAS":"CS","CHIHUAHUA":"CH",
    "CIUDAD DE MEXICO":"DF","CDMX":"DF","DISTRITO FEDERAL":"DF","DURANGO":"DG",
    "GUANAJUATO":"GT","GUERRERO":"GR","HIDALGO":"HG","JALISCO":"JC","MEXICO":"MC",
    "ESTADO DE MEXICO":"MC","EDO MEX":"MC","MICHOACAN":"MN","MORELOS":"MS",
    "NAYARIT":"NT","NUEVO LEON":"NL","OAXACA":"OC","PUEBLA":"PL","QUERETARO":"QT",
    "QUINTANA ROO":"QR","SAN LUIS POTOSI":"SP","SINALOA":"SL","SONORA":"SR",
    "TABASCO":"TC","TAMAULIPAS":"TS","TLAXCALA":"TL","VERACRUZ":"VZ",
    "YUCATAN":"YN","ZACATECAS":"ZS","EXTRANJERO":"NE"
  };
  return d[e] || e;
}

function curpValorCaracter(ch) {
  return "0123456789ABCDEFGHIJKLMNNOPQRSTUVWXYZ".indexOf(ch);
}

/**
 * Genera la CURP de una persona.
 * @customfunction GENERAR
 * @param {string} nombre Nombre(s) de la persona
 * @param {string} paterno Apellido paterno
 * @param {string} [materno] Apellido materno (opcional)
 * @param {string} fechaNacimiento Fecha AAAA-MM-DD
 * @param {string} sexo H, Hombre, M o Mujer
 * @param {string} estado Estado de nacimiento
 * @returns {string}
 */
function GENERAR(nombre, paterno, materno, fechaNacimiento, sexo, estado) {
  try {
    nombre  = limpiarTexto(nombre);
    paterno = limpiarTexto(paterno);
    materno = limpiarTexto(materno || "");
    sexo    = normalizarSexo(sexo);
    estado  = normalizarEstado(estado);
    if (!paterno) return "ERROR: Falta apellido paterno";
    if (!sexo)    return "ERROR: Sexo invalido";
    if (!estado)  return "ERROR: Estado invalido";
    const fecha = new Date(fechaNacimiento + "T12:00:00");
    if (isNaN(fecha)) return "ERROR: Fecha invalida";
    const yy = String(fecha.getFullYear()).slice(-2);
    const mm = String(fecha.getMonth()+1).padStart(2,"0");
    const dd = String(fecha.getDate()).padStart(2,"0");
    let base = paterno[0] + primeraVocalInterna(paterno) +
               (materno ? materno[0] : "X") + nombre[0] +
               yy + mm + dd + sexo + estado +
               consonanteInterna(paterno) +
               consonanteInterna(materno || "") +
               consonanteInterna(nombre);
    base += fecha.getFullYear() < 2000 ? "0" : "A";
    let suma = 0;
    for (let i = 0; i < base.length; i++)
      suma += curpValorCaracter(base[i]) * (19 - i);
    return base + (10 - (suma % 10)) % 10;
  } catch(e) { return "ERROR: " + e.message; }
}

function rfcFiltrar(nombre, paterno, materno) {
  const part = ["DE ","DEL ","LA ","LOS ","LAS ","Y ","MC ","MAC ","VON ","VAN "];
  for (const p of part) {
    nombre  = nombre.split(p).join("");
    paterno = paterno.split(p).join("");
    materno = materno.split(p).join("");
  }
  nombre  = nombre.replace(/[.,]/g,"").trim();
  paterno = paterno.replace(/[.,]/g,"").trim();
  materno = materno.replace(/[.,]/g,"").trim();
  if (nombre.includes(" "))
    for (const p of ["JOSE ","MARIA ","J ","MA "])
      nombre = nombre.split(p).join("");
  return { nombre:nombre.trim(), paterno:paterno.trim(), materno:materno.trim() };
}

function rfcProhibidas(rfc) {
  const p = ["BUEI","BUEY","CACA","CACO","CAGA","CAGO","CAKA","CAKO","COGE","COJA",
    "COJE","COJI","COJO","CULO","FETO","GUEY","JOTO","KACA","KACO","KAGA","KAGO",
    "KOGE","KOJO","KAKA","KULO","MAME","MAMO","MEAR","MEAS","MEON","MION","MOCO",
    "MULA","PEDA","PEDO","PENE","PUTA","PUTO","QULO","RATA","RUIN"];
  return p.includes(rfc.slice(0,4)) ? rfc.slice(0,3)+"X"+rfc.slice(4) : rfc;
}

function rfcHomoclave(nombre, paterno, materno) {
  const comp = paterno+" "+materno+" "+nombre;
  const chars = "123456789ABCDEFGHIJKLMNPQRSTUVWXYZ";
  let cad = "0";
  for (const ch of comp) {
    const c = ch.charCodeAt(0);
    if (ch===" "||ch==="-") cad+="00";
    else if (ch==="N"||ch==="U") cad+="10";
    else if (c>=65&&c<=73) cad+=String(c-54);
    else if (c>=74&&c<=82) cad+=String(c-53);
    else if (c>=83&&c<=90) cad+=String(c-51);
    else if (ch>="0"&&ch<="9") cad+=ch.padStart(2,"0");
  }
  let suma = 0;
  for (let i = 0; i < cad.length-1; i++)
    suma += parseInt(cad.slice(i,i+2)) * parseInt(cad[i+1]);
  const t = parseInt(String(suma).slice(-3))||0;
  return chars[Math.floor(t/34)] + chars[t%34];
}

function rfcDigito(rfc) {
  const chars = "0123456789ABCDEFGHIJKLMN&OPQRSTUVWXYZ ";
  let suma = 0;
  for (let i = 0; i < rfc.length; i++) {
    const idx = chars.indexOf(rfc[i]==" "?"*":rfc[i]);
    if (idx<0) return "0";
    suma += idx*(14-i);
  }
  const mod = suma%11;
  if (mod===0) return "0";
  const dv = 11-mod;
  return dv>9?"A":String(dv);
}

/**
 * Genera el RFC de una persona fisica.
 * @customfunction RFC
 * @param {string} nombre Nombre(s) de la persona
 * @param {string} paterno Apellido paterno
 * @param {string} [materno] Apellido materno (opcional)
 * @param {string} fechaNacimiento Fecha AAAA-MM-DD
 * @returns {string}
 */
function RFC(nombre, paterno, materno, fechaNacimiento) {
  try {
    nombre  = limpiarTexto(nombre);
    paterno = limpiarTexto(paterno);
    materno = limpiarTexto(materno||"");
    const fecha = new Date(fechaNacimiento+"T12:00:00");
    if (isNaN(fecha)) return "ERROR: Fecha invalida";
    const yy = String(fecha.getFullYear()).slice(-2);
    const mm = String(fecha.getMonth()+1).padStart(2,"0");
    const dd = String(fecha.getDate()).padStart(2,"0");
    const no = nombre, pa = paterno, ma = materno;
    const f  = rfcFiltrar(nombre, paterno, materno);
    let rfc;
    if (f.paterno && f.materno)
      rfc = f.paterno[0]+primeraVocalInterna(f.paterno)+(f.materno[0]||"X")+f.nombre[0]+yy+mm+dd;
    else
      rfc = (f.paterno+f.materno+f.nombre).slice(0,4)+yy+mm+dd;
    rfc = rfcProhibidas(rfc) + rfcHomoclave(no,pa,ma);
    return rfc + rfcDigito(rfc);
  } catch(e) { return "ERROR: "+e.message; }
}

/**
 * Extrae la fecha de nacimiento de una CURP.
 * @customfunction FECHA
 * @param {string} curp CURP de 18 caracteres
 * @returns {string}
 */
function FECHA(curp) {
  curp = limpiarTexto(curp);
  if (curp.length!==18) return "ERROR: CURP invalida";
  const anio = parseInt(curp.slice(4,6));
  const siglo = /[0-9]/.test(curp[16]) ? 1900 : 2000;
  return `${siglo+anio}-${curp.slice(6,8)}-${curp.slice(8,10)}`;
}

/**
 * Extrae el sexo de una CURP.
 * @customfunction SEXO
 * @param {string} curp CURP de 18 caracteres
 * @returns {string}
 */
function SEXO(curp) {
  curp = limpiarTexto(curp);
  if (curp.length!==18) return "ERROR: CURP invalida";
  return curp[10]==="H"?"Hombre":curp[10]==="M"?"Mujer":"ERROR";
}

/**
 * Calcula la edad actual desde una CURP.
 * @customfunction EDAD
 * @param {string} curp CURP de 18 caracteres
 * @returns {number}
 */
function EDAD(curp) {
  const f = FECHA(curp);
  if (f.startsWith("ERROR")) return f;
  const nac = new Date(f+"T12:00:00"), hoy = new Date();
  let edad = hoy.getFullYear()-nac.getFullYear();
  if (hoy < new Date(hoy.getFullYear(), nac.getMonth(), nac.getDate())) edad--;
  return edad;
}

/**
 * Extrae el estado de nacimiento de una CURP.
 * @customfunction ESTADO
 * @param {string} curp CURP de 18 caracteres
 * @returns {string}
 */
function ESTADO(curp) {
  curp = limpiarTexto(curp);
  if (curp.length!==18) return "ERROR: CURP invalida";
  const estados = {
    AS:"Aguascalientes",BC:"Baja California",BS:"Baja California Sur",
    CC:"Campeche",CS:"Chiapas",CH:"Chihuahua",CL:"Coahuila",CM:"Colima",
    DF:"Ciudad de Mexico",DG:"Durango",GT:"Guanajuato",GR:"Guerrero",
    HG:"Hidalgo",JC:"Jalisco",MC:"Mexico",MN:"Michoacan",MS:"Morelos",
    NT:"Nayarit",NL:"Nuevo Leon",OC:"Oaxaca",PL:"Puebla",QT:"Queretaro",
    QR:"Quintana Roo",SP:"San Luis Potosi",SL:"Sinaloa",SR:"Sonora",
    TC:"Tabasco",TS:"Tamaulipas",TL:"Tlaxcala",VZ:"Veracruz",
    YN:"Yucatan",ZS:"Zacatecas",NE:"Extranjero"
  };
  return estados[curp.slice(11,13)]||"Desconocido";
}

// Registro de funciones
CustomFunctions.associate("GENERAR", GENERAR);
CustomFunctions.associate("RFC",     RFC);
CustomFunctions.associate("FECHA",   FECHA);
CustomFunctions.associate("SEXO",    SEXO);
CustomFunctions.associate("EDAD",    EDAD);
CustomFunctions.associate("ESTADO",  ESTADO);
