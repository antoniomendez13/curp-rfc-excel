/* global CustomFunctions, Office */
"use strict";

function limpiarTexto(txt) {
  if (!txt) return "";
  var mapa = {"Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ü":"U","Ñ":"N","á":"A","é":"E","í":"I","ó":"O","ú":"U","ü":"U","ñ":"N"};
  var r = txt.toUpperCase().trim();
  for (var k in mapa) r = r.split(k).join(mapa[k]);
  return r;
}

function esVocal(l) { return "AEIOU".indexOf(l) >= 0; }

function primeraVocalInterna(txt) {
  for (var i = 1; i < txt.length; i++) if (esVocal(txt[i])) return txt[i];
  return "X";
}

function consonanteInterna(txt) {
  if (!txt) return "X";
  for (var i = 1; i < txt.length; i++)
    if (!esVocal(txt[i]) && txt[i] !== "N" && txt[i] !== " ") return txt[i];
  return "X";
}

function normalizarSexo(s) {
  s = limpiarTexto(s);
  if (s === "H" || s === "HOMBRE" || s === "MASCULINO") return "H";
  if (s === "M" || s === "MUJER"  || s === "FEMENINO")  return "M";
  return "";
}

function normalizarEstado(e) {
  e = limpiarTexto(e);
  var d = {
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
    var fecha = new Date(fechaNacimiento + "T12:00:00");
    if (isNaN(fecha.getTime())) return "ERROR: Fecha invalida";
    var yy = String(fecha.getFullYear()).slice(-2);
    var mm = String(fecha.getMonth() + 1).padStart(2, "0");
    var dd = String(fecha.getDate()).padStart(2, "0");
    var base = paterno[0] + primeraVocalInterna(paterno) +
               (materno ? materno[0] : "X") + nombre[0] +
               yy + mm + dd + sexo + estado +
               consonanteInterna(paterno) +
               consonanteInterna(materno || "") +
               consonanteInterna(nombre);
    base += fecha.getFullYear() < 2000 ? "0" : "A";
    var suma = 0;
    for (var i = 0; i < base.length; i++)
      suma += curpValorCaracter(base[i]) * (19 - i);
    return base + (10 - (suma % 10)) % 10;
  } catch (e) { return "ERROR: " + e.message; }
}

function rfcFiltrar(nombre, paterno, materno) {
  var part = ["DE ","DEL ","LA ","LOS ","LAS ","Y ","MC ","MAC ","VON ","VAN "];
  for (var i = 0; i < part.length; i++) {
    nombre  = nombre.split(part[i]).join("");
    paterno = paterno.split(part[i]).join("");
    materno = materno.split(part[i]).join("");
  }
  nombre  = nombre.replace(/[.,]/g, "").trim();
  paterno = paterno.replace(/[.,]/g, "").trim();
  materno = materno.replace(/[.,]/g, "").trim();
  if (nombre.indexOf(" ") >= 0) {
    var pref = ["JOSE ","MARIA ","J ","MA "];
    for (var j = 0; j < pref.length; j++) nombre = nombre.split(pref[j]).join("");
  }
  return { nombre: nombre.trim(), paterno: paterno.trim(), materno: materno.trim() };
}

function rfcProhibidas(rfc) {
  var p = ["BUEI","BUEY","CACA","CACO","CAGA","CAGO","CAKA","CAKO","COGE","COJA",
    "CEJE","COJI","COJO","CULO","FETO","GUEY","JOTO","KACA","KACO","KAGA","KAGO",
    "KOGE","KOJO","KAKA","KULO","MAME","MAMO","MEAR","MEAS","MEON","MION","MOCO",
    "MULA","PEDA","PEDO","PENE","PUTA","PUTO","QULO","RATA","RUIN"];
  return p.indexOf(rfc.slice(0, 4)) >= 0 ? rfc.slice(0, 3) + "X" + rfc.slice(4) : rfc;
}

function rfcHomoclave(nombre, paterno, materno) {
  var comp = paterno + " " + materno + " " + nombre;
  var chars = "123456789ABCDEFGHIJKLMNPQRSTUVWXYZ";
  var cad = "0";
  for (var i = 0; i < comp.length; i++) {
    var ch = comp[i];
    var c = comp.charCodeAt(i);
    if (ch === " " || ch === "-") cad += "00";
    else if (ch === "N" || ch === "U") cad += "10";
    else if (c >= 65 && c <= 73) cad += String(c - 54);
    else if (c >= 74 && c <= 82) cad += String(c - 53);
    else if (c >= 83 && c <= 90) cad += String(c - 51);
    else if (ch >= "0" && ch <= "9") cad += "0" + ch;
  }
  var suma = 0;
  for (var j = 0; j < cad.length - 1; j++)
    suma += parseInt(cad.slice(j, j + 2)) * parseInt(cad[j + 1]);
  var t = parseInt(String(suma).slice(-3)) || 0;
  return chars[Math.floor(t / 34)] + chars[t % 34];
}

function rfcDigito(rfc) {
  var chars = "0123456789ABCDEFGHIJKLMN&OPQRSTUVWXYZ ";
  var suma = 0;
  for (var i = 0; i < rfc.length; i++) {
    var idx = chars.indexOf(rfc[i] === " " ? "*" : rfc[i]);
    if (idx < 0) return "0";
    suma += idx * (14 - i);
  }
  var mod = suma % 11;
  if (mod === 0) return "0";
  var dv = 11 - mod;
  return dv > 9 ? "A" : String(dv);
}

function RFC(nombre, paterno, materno, fechaNacimiento) {
  try {
    var no = limpiarTexto(nombre);
    var pa = limpiarTexto(paterno);
    var ma = limpiarTexto(materno || "");
    var fecha = new Date(fechaNacimiento + "T12:00:00");
    if (isNaN(fecha.getTime())) return "ERROR: Fecha invalida";
    var yy = String(fecha.getFullYear()).slice(-2);
    var mm = String(fecha.getMonth() + 1).padStart(2, "0");
    var dd = String(fecha.getDate()).padStart(2, "0");
    var f = rfcFiltrar(no, pa, ma);
    var rfc;
    if (f.paterno && f.materno)
      rfc = f.paterno[0] + primeraVocalInterna(f.paterno) + (f.materno[0] || "X") + f.nombre[0] + yy + mm + dd;
    else
      rfc = (f.paterno + f.materno + f.nombre).slice(0, 4) + yy + mm + dd;
    rfc = rfcProhibidas(rfc) + rfcHomoclave(no, pa, ma);
    return rfc + rfcDigito(rfc);
  } catch (e) { return "ERROR: " + e.message; }
}

function FECHA(curp) {
  curp = limpiarTexto(curp);
  if (curp.length !== 18) return "ERROR: CURP invalida";
  var anio = parseInt(curp.slice(4, 6));
  var siglo = /[0-9]/.test(curp[16]) ? 1900 : 2000;
  return (siglo + anio) + "-" + curp.slice(6, 8) + "-" + curp.slice(8, 10);
}

function SEXO(curp) {
  curp = limpiarTexto(curp);
  if (curp.length !== 18) return "ERROR: CURP invalida";
  return curp[10] === "H" ? "Hombre" : curp[10] === "M" ? "Mujer" : "ERROR";
}

function EDAD(curp) {
  var f = FECHA(curp);
  if (f.indexOf("ERROR") === 0) return f;
  var nac = new Date(f + "T12:00:00");
  var hoy = new Date();
  var edad = hoy.getFullYear() - nac.getFullYear();
  if (hoy < new Date(hoy.getFullYear(), nac.getMonth(), nac.getDate())) edad--;
  return edad;
}

function ESTADO(curp) {
  curp = limpiarTexto(curp);
  if (curp.length !== 18) return "ERROR: CURP invalida";
  var estados = {
    AS:"Aguascalientes",BC:"Baja California",BS:"Baja California Sur",
    CC:"Campeche",CS:"Chiapas",CH:"Chihuahua",CL:"Coahuila",CM:"Colima",
    DF:"Ciudad de Mexico",DG:"Durango",GT:"Guanajuato",GR:"Guerrero",
    HG:"Hidalgo",JC:"Jalisco",MC:"Mexico",MN:"Michoacan",MS:"Morelos",
    NT:"Nayarit",NL:"Nuevo Leon",OC:"Oaxaca",PL:"Puebla",QT:"Queretaro",
    QR:"Quintana Roo",SP:"San Luis Potosi",SL:"Sinaloa",SR:"Sonora",
    TC:"Tabasco",TS:"Tamaulipas",TL:"Tlaxcala",VZ:"Veracruz",
    YN:"Yucatan",ZS:"Zacatecas",NE:"Extranjero"
  };
  return estados[curp.slice(11, 13)] || "Desconocido";
}

function registrar() {
  CustomFunctions.associate("GENERAR", GENERAR);
  CustomFunctions.associate("RFC",     RFC);
  CustomFunctions.associate("FECHA",   FECHA);
  CustomFunctions.associate("SEXO",    SEXO);
  CustomFunctions.associate("EDAD",    EDAD);
  CustomFunctions.associate("ESTADO",  ESTADO);
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(registrar);
} else {
  registrar();
}
