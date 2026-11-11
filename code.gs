/**
 * Proyecto Bethany - Salidas de fin de semana y solicitudes de mantenimiento (versión 9: instalación con código)
 * Perfiles: administración (admin), mamás (mama) y personal de mantenimiento (mant).
 * Aplicación web de Google Apps Script con Google Sheets como base de datos.
 *
 * Instalación nueva (una sola vez, con la cuenta de Google de Proyecto Bethany):
 *   1. Ejecutar instalar() desde el editor y autorizar los permisos.
 *      En el «Registro de ejecución» aparece un CÓDIGO DE INSTALACIÓN (solo lo ve quien tiene el proyecto).
 *   2. Implementar > Nueva implementación > Aplicación web
 *      (Ejecutar como: Yo; Quién tiene acceso: Cualquier usuario).
 *   3. Abrir la URL de la app, tocar «Configuración inicial» y crear la primera cuenta de administración
 *      con ese código. El código sirve una sola vez. Ninguna contraseña queda escrita en el código.
 *   Si se pierde el código antes de usarlo, ejecutar instalar() de nuevo para generar otro.
 *
 * Si ya estaba instalada una versión anterior:
 *   1. Reemplazar este archivo e Index.html.
 *   2. Ejecutar actualizar() una vez (no borra datos).
 *   3. Implementar > Administrar implementaciones > editar (lápiz) > Versión: Nueva versión > Implementar.
 */

// ---------------------------------------------------------------------------
// Configuración inicial
// ---------------------------------------------------------------------------
const TZ = 'America/Costa_Rica';
const VERSION_ESTRUCTURA = '8';   // si cambia, actualizar() se ejecuta sola en el siguiente uso de la app
const SESION_SEG = 21600;          // 6 horas, máximo permitido por CacheService
const ITER_HASH = 1000;            // iteraciones del hash de contraseñas
const MAX_INTENTOS = 5;            // intentos fallidos antes de bloquear el ingreso
const BLOQUEO_SEG = 900;           // 15 minutos de bloqueo
const NOMBRES_DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const GENEROS = ['Femenino', 'Masculino'];
const SEMANAS_CUATRI = 15;          // semanas del cuatrimestre (la fecha de fin cae en la última)
const TOKEN_NORMAL = 'Normal', TOKEN_EXC = 'Excepción';

// Relleno de las filas en los Excel según el estado de la solicitud
const COLOR_ESTADO = { Aprobada: '#c6efce', Rechazada: '#ffc7ce', Pendiente: '#d9d9d9', Anulada: '#ffeb9c' };

// Mantenimiento y reparaciones (boleta de solicitud de mantenimiento)
const ESTADOS_MANT = ['Pendiente', 'En proceso', 'En espera de materiales o repuestos', 'Finalizada'];
const PRIORIDADES_MANT = ['Alta', 'Media', 'Baja'];
const DESDE_MANT = ['Hoy', 'Hace varios días', 'Hace varias semanas', 'No estoy segura'];
const COLOR_MANT = { Pendiente: '#d9d9d9', 'En proceso': '#ddebf7', 'En espera de materiales o repuestos': '#ffeb9c', Finalizada: '#c6efce' };
const COLS_MANT = ['ID', 'Enviada', 'Cedula_mama', 'Nombre', 'Telefono', 'Correo', 'Apartamento', 'Descripcion', 'Fecha_aproximada',
                   'Desde_cuando', 'Riesgo', 'Riesgo_detalle', 'Atencion_inmediata', 'Info_adicional', 'Firma', 'Estado', 'Prioridad',
                   'Fecha_recepcion', 'Evaluacion', 'Trabajo', 'Materiales', 'Responsable', 'Fecha_atencion', 'Observaciones',
                   'Fecha_finalizacion', 'Firma_responsable', 'Actualizada_por', 'Actualizada'];
// Registro permanente de boletas: nunca se borran filas, aunque la boleta se borre de la app.
const COLS_REG_MANT = COLS_MANT.concat(['En_la_app', 'Borrada_por', 'Fecha_borrado']);
const HOJAS_REG_MANT = { RegPendiente: 'Pendiente', RegProceso: 'En proceso', RegEspera: 'En espera de materiales o repuestos', RegFinalizada: 'Finalizada' };
const MAX_FIRMA = 48000;           // caracteres de la imagen de la firma (una celda de Sheets admite 50 000)

const COLS_SOL = ['ID', 'Enviada', 'Cedula_mama', 'Nombre', 'Correo', 'Telefono', 'Salida', 'Hora_salida',
                  'Regreso', 'Hora_regreso', 'Dias', 'Token', 'Semana', 'Comentario', 'Estado', 'Revisada_por', 'Fecha_revision', 'Nota_admin'];
// Columnas de versiones anteriores que cambiaron de nombre
const REN_SOL = { ID_Mama: 'Cedula_mama', Aprobada_por: 'Revisada_por', Fecha_aprobacion: 'Fecha_revision' };

// Columnas que siempre se guardan como texto (cédulas, teléfonos, horas e identificadores)
const TEXTO_COLS = ['ID', 'Usuario', 'Cedula', 'Cedula_mama', 'Cedula_hijo', 'ID_Usuario', 'Telefono', 'Hora_salida', 'Hora_regreso', 'Apartamento'];

const LIBROS = {
  ID_SISTEMA: 'Bethany - Sistema (NO COMPARTIR)',
  ID_SOLICITUDES: 'Bethany - Solicitudes',
  ID_APROBADAS: 'Bethany - Solicitudes aprobadas',
  ID_RESIDENTES: 'Bethany - Residentes',
  ID_MANTENIMIENTO: 'Bethany - Mantenimiento',
  ID_REG_MANT: 'Bethany - Registro de mantenimiento'
};

// Tablas: libro, hoja y columnas. El orden de declaración define el orden de las hojas en cada libro.
const TABLAS = {
  Usuarios: { libro: 'ID_SISTEMA', hoja: 'Usuarios',
              cols: ['ID', 'Usuario', 'Cedula', 'Nombre', 'Correo', 'Avisos_correo', 'Telefono', 'Rol', 'Estado', 'Hash', 'Salt', 'Creado', 'Consentimiento'] },
  Config: { libro: 'ID_SISTEMA', hoja: 'Config', cols: ['Clave', 'Valor', 'Descripcion'] },
  Avisos: { libro: 'ID_SISTEMA', hoja: 'Avisos', cols: ['ID', 'Fecha', 'ID_Usuario', 'Mensaje', 'Leido'] },
  Firmas: { libro: 'ID_SISTEMA', hoja: 'Firmas', cols: ['ID', 'Tipo', 'Imagen', 'Fecha'] },
  Todas: { libro: 'ID_SOLICITUDES', hoja: 'Todas', cols: COLS_SOL, ren: REN_SOL, color: true },
  SolAprobadas: { libro: 'ID_SOLICITUDES', hoja: 'Aprobadas', cols: COLS_SOL, ren: REN_SOL, color: true },
  SolRechazadas: { libro: 'ID_SOLICITUDES', hoja: 'Rechazadas', cols: COLS_SOL, ren: REN_SOL, color: true },
  SolPendientes: { libro: 'ID_SOLICITUDES', hoja: 'Pendientes', cols: COLS_SOL, ren: REN_SOL, color: true },
  Registro: { libro: 'ID_APROBADAS', hoja: 'Aprobadas', cols: COLS_SOL, ren: REN_SOL, color: true },
  Mamas: { libro: 'ID_RESIDENTES', hoja: 'Mamás', ren: { ID: 'Cedula' }, edad: 'Fecha_nacimiento',
           cols: ['Cedula', 'Nombre', 'Correo', 'Telefono', 'Fecha_nacimiento', 'Edad', 'Universidad', 'Carrera', 'Porcentaje_beca',
                  'Cuatrimestres', 'Cantidad_hijos', 'Hijos', 'Estado_cuenta', 'Creado', 'Actualizado'] },
  Hijos: { libro: 'ID_RESIDENTES', hoja: 'Hijos', ren: { ID_Mama: 'Cedula_mama' }, edad: 'Fecha_nacimiento',
           cols: ['Cedula_mama', 'Nombre_mama', 'Nombre_hijo', 'Cedula_hijo', 'Fecha_nacimiento', 'Edad', 'Genero'] },
  Mantenimiento: { libro: 'ID_MANTENIMIENTO', hoja: 'Solicitudes', cols: COLS_MANT, color: true, colores: COLOR_MANT },
  RegTodas: { libro: 'ID_REG_MANT', hoja: 'Todas (por estado)', cols: COLS_REG_MANT, color: true, colores: COLOR_MANT },
  RegPendiente: { libro: 'ID_REG_MANT', hoja: 'Pendientes', cols: COLS_REG_MANT, color: true, colores: COLOR_MANT },
  RegProceso: { libro: 'ID_REG_MANT', hoja: 'En proceso', cols: COLS_REG_MANT, color: true, colores: COLOR_MANT },
  RegEspera: { libro: 'ID_REG_MANT', hoja: 'En espera de materiales', cols: COLS_REG_MANT, color: true, colores: COLOR_MANT },
  RegFinalizada: { libro: 'ID_REG_MANT', hoja: 'Finalizadas', cols: COLS_REG_MANT, color: true, colores: COLOR_MANT }
};
const TABLAS_SOL = ['Todas', 'SolAprobadas', 'SolRechazadas', 'SolPendientes', 'Registro'];

// ---------------------------------------------------------------------------
// Entrada de la aplicación web
// ---------------------------------------------------------------------------
function doGet() {
  // ALLOWALL permite mostrar la app dentro de la página de GitHub Pages (iframe a pantalla completa).
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Proyecto Bethany | Salidas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------------------
// Instalación y actualización
// ---------------------------------------------------------------------------
function instalar() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ID_SISTEMA') && hayAdmin_()) {
    throw new Error('El sistema ya está instalado y tiene administración. Si cambiaste el código, solo hacé una nueva versión de la implementación.');
  }
  actualizar();
  const letras = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let codigo = '';
  for (let i = 0; i < 12; i++) codigo += letras.charAt(Math.floor(Math.random() * letras.length));
  codigo = codigo.slice(0, 4) + '-' + codigo.slice(4, 8) + '-' + codigo.slice(8);
  props.setProperty('CODIGO_INICIAL', codigo);
  const msg = 'Instalación lista. CÓDIGO DE INSTALACIÓN: ' + codigo +
              '  →  Usalo en la app («Configuración inicial») para crear la primera cuenta de administración. Sirve una sola vez.';
  console.log(msg);
  Logger.log(msg);
  return msg;
}

function hayAdmin_() {
  return filas_('Usuarios').some(function (u) { return u.Rol === 'admin'; });
}

/** La pantalla de ingreso pregunta si falta crear la primera cuenta de administración. */
function estadoInstalacion() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ID_SISTEMA')) return { instalado: false, necesitaAdmin: false };
  asegurarEstructura_();
  return { instalado: true, necesitaAdmin: !hayAdmin_() && !!props.getProperty('CODIGO_INICIAL') };
}

/** Crea la primera cuenta de administración con el código de instalación (una sola vez). */
function crearPrimerAdmin(codigo, d) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ID_SISTEMA')) throw new Error('Primero hay que ejecutar instalar() desde el editor de Apps Script.');
  const cache = CacheService.getScriptCache();
  const intentos = Number(cache.get('f_codigo_inicial') || 0);
  if (intentos >= MAX_INTENTOS) throw new Error('Demasiados intentos con un código incorrecto. Esperá 15 minutos.');
  d = d || {};
  const datos = validarCuenta_(d);
  return conBloqueo_(function () {
    const guardado = props.getProperty('CODIGO_INICIAL');
    if (!guardado || hayAdmin_()) throw new Error('La configuración inicial ya se hizo. Ingresá con tu cuenta de administración.');
    const limpio = function (c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); };
    if (limpio(codigo) !== limpio(guardado)) {
      cache.put('f_codigo_inicial', String(intentos + 1), BLOQUEO_SEG);
      throw new Error('El código de instalación no es correcto. Copialo del «Registro de ejecución» de instalar().');
    }
    revisarIdentificadores_(filas_('Usuarios'), null, datos);
    const id = 'U-' + Utilities.getUuid().slice(0, 8);
    crearUsuario_(id, { nombre: d.nombre, correo: datos.correo, telefono: d.telefono || '', usuario: datos.usuario, cedula: datos.cedula,
                        consentimiento: true }, d.clave, 'admin', 'Activa');
    props.deleteProperty('CODIGO_INICIAL');
    cache.remove('f_codigo_inicial');
    const token = Utilities.getUuid();
    cache.put('s_' + token, JSON.stringify({ id: id }), SESION_SEG);
    return { token: token, mensaje: 'Cuenta de administración creada. ¡Bienvenida o bienvenido!' };
  });
}

function actualizar() {
  const props = PropertiesService.getScriptProperties();
  conBloqueo_(function () {
    Object.keys(LIBROS).forEach(function (prop) {
      let ss;
      let nuevo = false;
      if (props.getProperty(prop)) {
        ss = SpreadsheetApp.openById(props.getProperty(prop));
      } else {
        ss = SpreadsheetApp.create(LIBROS[prop]);
        ss.setSpreadsheetTimeZone(TZ);
        props.setProperty(prop, ss.getId());
        nuevo = true;
      }
      if (prop === 'ID_SOLICITUDES' && !ss.getSheetByName('Todas') && ss.getSheetByName('Solicitudes')) {
        ss.getSheetByName('Solicitudes').setName('Todas');
      }
      Object.keys(TABLAS).filter(function (k) { return TABLAS[k].libro === prop; }).forEach(function (k, i) {
        const t = TABLAS[k];
        let sh = ss.getSheetByName(t.hoja);
        if (!sh) sh = (nuevo && i === 0) ? ss.getSheets()[0].setName(t.hoja) : ss.insertSheet(t.hoja);
        ordenarColumnas_(sh, t);
        ss.setActiveSheet(sh);
        ss.moveActiveSheet(i + 1);
      });
      Logger.log('%s: %s', LIBROS[prop], ss.getUrl());
    });
    cache_ = {};

    // Versión 6: los días por mes se reemplazan por tokens (una salida = un token, sin límite de días).
    borrarFilas_('Config', filas_('Config').filter(function (r) { return r.Clave === 'LIMITE_MENSUAL'; }).map(function (r) { return r._fila; }));
    const cfg = filas_('Config');
    [['TOKENS_MES', 2, 'Tokens (salidas) por mamá en cada mes. Cada salida usa un token, sin importar cuántos días dure'],
     ['TOKENS_EXCEPCION', 1, 'Tokens de excepción por mamá en cada mes, para cuando ya usó todos sus tokens'],
     ['DIAS_SALIDA', 'viernes,sábado', 'Días de la semana en que puede iniciar una salida'],
     ['CORREOS_MAMAS', 'Sí', 'Enviar correo a las mamás cuando se aprueba, rechaza o anula una solicitud'],
     ['CUATRI_INICIO', '', 'Fecha en que empieza la semana 1 del cuatrimestre'],
     ['CUATRI_FIN', '', 'Fecha en que termina el cuatrimestre (semana ' + SEMANAS_CUATRI + ')']].forEach(function (c) {
      if (!cfg.some(function (r) { return r.Clave === c[0]; })) agregar_('Config', { Clave: c[0], Valor: c[1], Descripcion: c[2] });
    });

    // Versión 6: las solicitudes anteriores usaron un token normal.
    TABLAS_SOL.forEach(function (k) {
      filas_(k).forEach(function (r) { if (!r.Token) actualizarCelda_(k, r._fila, 'Token', TOKEN_NORMAL); });
    });
    // Versión 6: la cédula de las mamás también queda en la columna Cedula de Usuarios.
    filas_('Usuarios').forEach(function (u) {
      if (u.Rol === 'mama' && cedulaValida_(u.ID) && u.Cedula !== u.ID) actualizarCelda_('Usuarios', u._fila, 'Cedula', u.ID);
    });

    // Versión 1: las aprobadas solo estaban en el libro de aprobadas.
    filas_('Registro').forEach(function (r) {
      if (!r.Estado) actualizarCelda_('Registro', r._fila, 'Estado', 'Aprobada');
    });
    const idsTodas = filas_('Todas').map(function (r) { return r.ID; });
    filas_('Registro').forEach(function (r) {
      if (idsTodas.indexOf(r.ID) === -1) agregar_('Todas', copiaSol_(r, { Estado: 'Aprobada' }));
    });

    // Las hojas Aprobadas, Rechazadas y Pendientes reflejan el estado de la hoja Todas.
    const todas = filas_('Todas');
    [['SolAprobadas', 'Aprobada'], ['SolRechazadas', 'Rechazada'], ['SolPendientes', 'Pendiente']].forEach(function (par) {
      const actuales = filas_(par[0]);
      const debe = todas.filter(function (r) { return r.Estado === par[1]; });
      const idsDebe = debe.map(function (r) { return r.ID; });
      borrarFilas_(par[0], actuales.filter(function (r) { return idsDebe.indexOf(r.ID) === -1; }).map(function (r) { return r._fila; }));
      const idsActuales = actuales.map(function (r) { return r.ID; });
      debe.forEach(function (r) { if (idsActuales.indexOf(r.ID) === -1) agregar_(par[0], copiaSol_(r)); });
    });

    // Reparación: solicitudes que quedaron con un ID viejo de la mamá (por ejemplo, si un guardado de perfil
    // falló a medias antes de esta versión) se reasignan usando el correo.
    const usuarios = filas_('Usuarios');
    const ids = usuarios.map(function (u) { return u.ID; });
    const idPorCorreo = {};
    usuarios.forEach(function (u) { if (u.Rol === 'mama') idPorCorreo[u.Correo] = u.ID; });
    TABLAS_SOL.forEach(function (k) {
      filas_(k).forEach(function (r) {
        if (ids.indexOf(r.Cedula_mama) === -1 && idPorCorreo[r.Correo]) actualizarCelda_(k, r._fila, 'Cedula_mama', idPorCorreo[r.Correo]);
      });
    });

    // Cada mamá debe tener su fila en el libro de residentes.
    const idsMamas = filas_('Mamas').map(function (r) { return r.Cedula; });
    filas_('Usuarios').forEach(function (u) {
      if (u.Rol === 'mama' && idsMamas.indexOf(u.ID) === -1) {
        agregar_('Mamas', { Cedula: u.ID, Nombre: u.Nombre, Correo: u.Correo, Telefono: u.Telefono,
                            Estado_cuenta: u.Estado, Creado: fechaHora_(u.Creado) });
      }
    });

    TABLAS_SOL.forEach(recolorear_);
    recolorear_('Mantenimiento');
    sincronizarRegistroMant_();
    refrescarEdades_('Mamas');
    refrescarEdades_('Hijos');
    props.setProperty('VERSION_ESTRUCTURA', VERSION_ESTRUCTURA);
  });
  Logger.log('Actualización lista.');
}

/** Deja las columnas de la hoja en el orden indicado sin perder datos. */
function ordenarColumnas_(sh, t) {
  const cols = t.cols, ren = t.ren || {};
  const ultCol = sh.getLastColumn();
  const ultFila = sh.getLastRow();
  const datos = ultCol && ultFila ? sh.getRange(1, 1, ultFila, ultCol).getValues() : [];
  const enc = datos.length ? datos[0].map(function (c) { return ren[String(c)] || String(c); }) : [];
  const extra = enc.filter(function (c) { return c && cols.indexOf(c) === -1; });
  const final = cols.concat(extra);
  const original = datos.length ? datos[0].map(String) : [];
  const igual = original.length === final.length && original.every(function (c, i) { return c === final[i]; });
  if (!igual) {
    const filas = datos.slice(1).map(function (r) {
      return final.map(function (c) {
        const j = enc.indexOf(c);
        return valorCelda_(c, j === -1 ? '' : r[j]);
      });
    });
    sh.clearContents();
    sh.getRange(1, 1, 1, final.length).setValues([final]);
    if (filas.length) sh.getRange(2, 1, filas.length, final.length).setValues(filas);
  }
  sh.getRange(1, 1, 1, final.length).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/** Si se pegó código nuevo y no se ejecutó actualizar(), la estructura se actualiza sola. */
function asegurarEstructura_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ID_SISTEMA') && props.getProperty('VERSION_ESTRUCTURA') !== VERSION_ESTRUCTURA) actualizar();
  if (props.getProperty('ID_SISTEMA')) edadesDelDia_();
}

// ---------------------------------------------------------------------------
// Acceso a datos
// ---------------------------------------------------------------------------
let cache_ = {};

function hoja_(k) {
  const t = TABLAS[k];
  const id = PropertiesService.getScriptProperties().getProperty(t.libro);
  if (!id) throw new Error('El sistema no está instalado. Ejecutá instalar() o actualizar() desde el editor.');
  const sh = SpreadsheetApp.openById(id).getSheetByName(t.hoja);
  if (!sh) throw new Error('Falta la hoja "' + t.hoja + '". Ejecutá actualizar() desde el editor.');
  return sh;
}

function encabezado_(k) {
  if (!cache_[k]) {
    const sh = hoja_(k);
    cache_[k] = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  }
  return cache_[k];
}

/** Prepara un valor para escribirlo en una celda sin que Sheets lo convierta. */
function valorCelda_(c, v) {
  if (v === undefined || v === null) return '';
  if (TEXTO_COLS.indexOf(c) !== -1 && v !== '') {
    const s = (c === 'Hora_salida' || c === 'Hora_regreso') ? hora_(v) : String(v).replace(/^'/, '');
    return "'" + s;
  }
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

function filas_(k) {
  const sh = hoja_(k);
  if (sh.getLastRow() < 2) return [];
  const valores = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const enc = valores[0];
  return valores.slice(1).map(function (r, i) {
    const o = { _fila: i + 2 };
    enc.forEach(function (c, j) {
      let v = r[j];
      if (TEXTO_COLS.indexOf(c) !== -1) v = (c === 'Hora_salida' || c === 'Hora_regreso') ? hora_(v) : String(v);
      o[c] = v;
    });
    return o;
  });
}

function agregar_(k, obj) {
  const sh = hoja_(k);
  const enc = encabezado_(k);
  sh.appendRow(enc.map(function (c) { return valorCelda_(c, obj[c]); }));
  const fila = sh.getLastRow();
  if (TABLAS[k].color) colorear_(k, fila, obj.Estado);
  return fila;
}

function actualizarCelda_(k, fila, columna, valor) {
  hoja_(k).getRange(fila, encabezado_(k).indexOf(columna) + 1).setValue(valorCelda_(columna, valor));
  if (columna === 'Estado' && TABLAS[k].color) colorear_(k, fila, valor);
}

function actualizarFila_(k, fila, obj) {
  Object.keys(obj).forEach(function (c) { actualizarCelda_(k, fila, c, obj[c]); });
}

function borrarFilas_(k, filas) {
  const sh = hoja_(k);
  filas.slice().sort(function (a, b) { return b - a; }).forEach(function (f) { sh.deleteRow(f); });
}

function colorear_(k, fila, estado) {
  hoja_(k).getRange(fila, 1, 1, encabezado_(k).length).setBackground((TABLAS[k].colores || COLOR_ESTADO)[estado] || null);
}

function recolorear_(k) {
  const datos = filas_(k);
  if (!datos.length) return;
  const ancho = encabezado_(k).length;
  hoja_(k).getRange(2, 1, datos.length, ancho).setBackgrounds(datos.map(function (r) {
    const color = (TABLAS[k].colores || COLOR_ESTADO)[r.Estado] || null;
    return Array.apply(null, Array(ancho)).map(function () { return color; });
  }));
}

/** La columna Edad usa una fórmula, así la edad siempre está al día en Sheets y en el Excel. */
/**
 * La columna Edad se guarda como número calculado por el script, no como fórmula, para que nunca muestre
 * #ERROR! (las fórmulas escritas desde Apps Script pueden fallar según la configuración regional de la hoja).
 * Se recalcula al guardar un perfil, al ejecutar actualizar() y una vez al día cuando alguien usa la app.
 */
function refrescarEdades_(k) {
  const sh = hoja_(k);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const enc = encabezado_(k);
  const colFecha = enc.indexOf(TABLAS[k].edad) + 1, colEdad = enc.indexOf('Edad') + 1;
  if (colFecha < 1 || colEdad < 1) return;
  const hoy = hoyYmd_();
  const fechas = sh.getRange(2, colFecha, n, 1).getValues();
  sh.getRange(2, colEdad, n, 1).setValues(fechas.map(function (f) {
    const ymd = ymd_(f[0]);
    return [fechaValida_(ymd) && ymd <= hoy ? edad_(ymd, hoy) : ''];
  }));
}

/** Recalcula las edades una vez por día (cumpleaños), sin necesidad de activadores. */
function edadesDelDia_() {
  const props = PropertiesService.getScriptProperties();
  const hoy = hoyYmd_();
  if (props.getProperty('EDADES_DIA') === hoy) return;
  try {
    refrescarEdades_('Mamas');
    refrescarEdades_('Hijos');
    props.setProperty('EDADES_DIA', hoy);
  } catch (e) {
    Logger.log('No se pudieron recalcular las edades: %s', e);
  }
}

function conBloqueo_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function config_() {
  const c = {};
  filas_('Config').forEach(function (r) { c[r.Clave] = r.Valor; });
  const dias = String(c.DIAS_SALIDA || '').split(',').map(function (s) {
    return NOMBRES_DIAS.indexOf(s.trim().toLowerCase());
  }).filter(function (n) { return n >= 0; });
  const num = function (v, def) { return v === undefined || v === '' || isNaN(Number(v)) ? def : Number(v); };
  const inicio = ymd_(c.CUATRI_INICIO), fin = ymd_(c.CUATRI_FIN);
  return { tokens: num(c.TOKENS_MES, 2), excepcion: num(c.TOKENS_EXCEPCION, 1), diasSalida: dias,
           correosMamas: c.CORREOS_MAMAS !== 'No',
           inicio: fechaValida_(inicio) ? inicio : '', fin: fechaValida_(fin) ? fin : '' };
}

function escribirConfig_(clave, valor) {
  const r = filas_('Config').find(function (x) { return x.Clave === clave; });
  if (r) actualizarCelda_('Config', r._fila, 'Valor', valor);
  else agregar_('Config', { Clave: clave, Valor: valor, Descripcion: '' });
}

// ---------------------------------------------------------------------------
// Semanas del cuatrimestre
// ---------------------------------------------------------------------------
/**
 * Semana del cuatrimestre para una fecha. La semana 1 empieza en la fecha de inicio y cada semana dura 7 días.
 * Desde el inicio hasta la fecha de fin se cuenta como máximo la semana 15 (por ejemplo, si el cuatrimestre
 * empieza el 7 de septiembre, el 30 de octubre es la semana 8).
 */
function semanaCuatri_(cfg, fecha) {
  if (!cfg.inicio || !cfg.fin || !fechaValida_(fecha)) return { estado: 'sin_fechas', inicio: cfg.inicio, fin: cfg.fin };
  const base = { inicio: cfg.inicio, fin: cfg.fin, total: SEMANAS_CUATRI };
  if (fecha < cfg.inicio) return Object.assign(base, { estado: 'antes', faltan: Math.round((utc_(cfg.inicio) - utc_(fecha)) / 86400000) });
  if (fecha > cfg.fin) return Object.assign(base, { estado: 'terminado' });
  const n = Math.floor((utc_(fecha) - utc_(cfg.inicio)) / (7 * 86400000)) + 1;
  return Object.assign(base, { estado: 'en_curso', semana: Math.min(n, SEMANAS_CUATRI) });
}

function numeroSemana_(cfg, fecha) {
  const s = semanaCuatri_(cfg, fecha);
  return s.estado === 'en_curso' ? s.semana : '';
}

// ---------------------------------------------------------------------------
// Fechas, horas y conteo de días
// ---------------------------------------------------------------------------
function esFecha_(v) { return Object.prototype.toString.call(v) === '[object Date]'; }

function ymd_(v) {
  if (esFecha_(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

function fechaHora_(v) {
  if (esFecha_(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm');
  return String(v || '');
}

function hora_(v) {
  if (esFecha_(v)) return Utilities.formatDate(v, TZ, 'HH:mm');
  return String(v || '').replace(/^'/, '').slice(0, 5);
}

function horaValida_(h) { return /^([01]\d|2[0-3]):[0-5]\d$/.test(String(h || '')); }
function fechaValida_(f) { return /^\d{4}-\d{2}-\d{2}$/.test(String(f || '')) && !isNaN(utc_(f)); }

function aFecha_(ymd) {
  const p = ymd.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function hoyYmd_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }

function utc_(ymd) {
  const p = String(ymd).split('-').map(Number);
  return Date.UTC(p[0], p[1] - 1, p[2]);
}

/** Días (noches fuera) por mes. Cada noche cuenta en el mes en que empieza. */
function nochesPorMes_(salida, regreso) {
  const res = {};
  for (let t = utc_(salida); t < utc_(regreso); t += 86400000) {
    const d = new Date(t);
    const mes = d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2);
    res[mes] = (res[mes] || 0) + 1;
  }
  return res;
}

function diaSemana_(ymd) { return new Date(utc_(ymd)).getUTCDay(); }

function edad_(nacimiento, hoy) {
  const n = nacimiento.split('-').map(Number), h = hoy.split('-').map(Number);
  let e = h[0] - n[0];
  if (h[1] < n[1] || (h[1] === n[1] && h[2] < n[2])) e--;
  return e;
}

// ---------------------------------------------------------------------------
// Solicitudes: formato y reglas
// ---------------------------------------------------------------------------
function mapaSolicitud_(r) {
  return {
    id: r.ID, enviada: fechaHora_(r.Enviada), idMama: r.Cedula_mama, cedula: r.Cedula_mama, nombre: String(r.Nombre || ''),
    correo: String(r.Correo || ''), telefono: String(r.Telefono || ''),
    salida: ymd_(r.Salida), horaSalida: hora_(r.Hora_salida), regreso: ymd_(r.Regreso), horaRegreso: hora_(r.Hora_regreso),
    dias: Number(r.Dias) || 0, token: r.Token === TOKEN_EXC ? TOKEN_EXC : TOKEN_NORMAL,
    semana: r.Semana === '' || r.Semana == null ? '' : Number(r.Semana), comentario: String(r.Comentario || ''), estado: String(r.Estado || ''),
    revisadaPor: String(r.Revisada_por || ''), fechaRevision: fechaHora_(r.Fecha_revision), notaAdmin: String(r.Nota_admin || '')
  };
}

/** Copia una fila de solicitud (leída de una hoja) para escribirla en otra, con cambios opcionales. */
function copiaSol_(r, cambios) {
  const o = {};
  COLS_SOL.forEach(function (c) { o[c] = r[c] === undefined ? '' : r[c]; });
  Object.keys(cambios || {}).forEach(function (c) { o[c] = cambios[c]; });
  return o;
}

function quitarDe_(k, id) {
  borrarFilas_(k, filas_(k).filter(function (x) { return x.ID === id; }).map(function (x) { return x._fila; }));
}

function salidasDe_(idMama, incluirPendientes) {
  return filas_('Todas').filter(function (r) {
    return r.Cedula_mama === idMama && (r.Estado === 'Aprobada' || (incluirPendientes && r.Estado === 'Pendiente'));
  }).map(function (r) {
    return { id: r.ID, salida: ymd_(r.Salida), regreso: ymd_(r.Regreso), estado: r.Estado,
             token: r.Token === TOKEN_EXC ? TOKEN_EXC : TOKEN_NORMAL };
  });
}

/** Tokens usados por mes (el mes en que inicia la salida): { '2026-11': { normal: 1, excepcion: 0 } }. */
function tokensPorMes_(salidas) {
  const usos = {};
  salidas.forEach(function (s) {
    const m = String(s.salida).slice(0, 7);
    usos[m] = usos[m] || { normal: 0, excepcion: 0 };
    if (s.token === TOKEN_EXC) usos[m].excepcion++; else usos[m].normal++;
  });
  return usos;
}

function nombreMes_(m) {
  const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const p = String(m).split('-');
  return meses[Number(p[1]) - 1] + ' ' + p[0];
}

/**
 * Valida una salida y decide qué token usa. Cada salida usa un token del mes en que inicia, sin límite de días.
 * opciones.usarExcepcion = la mamá eligió usar su token de excepción (solo si ya no le quedan tokens normales).
 * opciones.alAprobar = true revisa solo contra las aprobadas, con el token ya elegido (opciones.token),
 * y omite las reglas de fecha mínima y día de inicio.
 */
function validarSalida_(idMama, salida, regreso, excluirId, opciones) {
  opciones = opciones || {};
  if (!fechaValida_(salida) || !fechaValida_(regreso)) return { error: 'Fechas inválidas.' };
  if (regreso <= salida) return { error: 'La fecha de regreso debe ser posterior a la de salida.' };
  const cfg = config_();
  if (!opciones.alAprobar) {
    if (salida < hoyYmd_()) return { error: 'La fecha de salida ya pasó.' };
    if (cfg.diasSalida.length && cfg.diasSalida.indexOf(diaSemana_(salida)) === -1) {
      return { error: 'La salida debe iniciar un ' + cfg.diasSalida.map(function (n) { return NOMBRES_DIAS[n]; }).join(' o ') + '.' };
    }
  }
  const otras = salidasDe_(idMama, !opciones.alAprobar).filter(function (s) { return s.id !== excluirId; });
  const choque = otras.find(function (s) { return salida < s.regreso && s.salida < regreso; });
  if (choque) return { error: 'Esas fechas se cruzan con otra salida (' + choque.estado.toLowerCase() + ') del ' + choque.salida + ' al ' + choque.regreso + '.' };
  let dias = 0;
  const pedidas = nochesPorMes_(salida, regreso);
  Object.keys(pedidas).forEach(function (m) { dias += pedidas[m]; });
  const mes = salida.slice(0, 7);
  const uso = tokensPorMes_(otras)[mes] || { normal: 0, excepcion: 0 };
  const quedanNormales = Math.max(cfg.tokens - uso.normal, 0);
  const quedanExc = Math.max(cfg.excepcion - uso.excepcion, 0);
  if (opciones.alAprobar) {
    if (opciones.token === TOKEN_EXC) {
      if (!quedanExc) return { error: 'ya tiene aprobados ' + uso.excepcion + ' de ' + cfg.excepcion + ' tokens de excepción en ' + nombreMes_(mes) + '.' };
    } else if (!quedanNormales) {
      return { error: 'ya tiene aprobados ' + uso.normal + ' de ' + cfg.tokens + ' tokens en ' + nombreMes_(mes) + '.' };
    }
    return { dias: dias, token: opciones.token === TOKEN_EXC ? TOKEN_EXC : TOKEN_NORMAL };
  }
  if (quedanNormales) return { dias: dias, token: TOKEN_NORMAL };
  const tusTokens = cfg.tokens === 1 ? 'tu token' : 'tus ' + cfg.tokens + ' tokens';
  if (!quedanExc) {
    return { error: 'Ya usaste ' + tusTokens + (cfg.excepcion ? ' y tu token de excepción' : '') + ' de ' + nombreMes_(mes) + '.' };
  }
  if (!opciones.usarExcepcion) {
    return { error: 'Ya usaste ' + tusTokens + ' de ' + nombreMes_(mes) + '. Si necesitás salir, marcá la opción para usar tu token de excepción.',
             pideExcepcion: true };
  }
  return { dias: dias, token: TOKEN_EXC };
}

// ---------------------------------------------------------------------------
// Cuentas, perfil y sesiones
// ---------------------------------------------------------------------------
function texto_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 500); }

/** Nombre de usuario: 3 a 30 caracteres (letras, números, punto, guion o guion bajo) y al menos una letra. */
function normalizarUsuario_(u) { return String(u == null ? '' : u).trim().toLowerCase(); }
function usuarioValido_(u) { return /^[a-z0-9._-]{3,30}$/.test(u) && /[a-z]/.test(u); }

function correoValido_(c) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(c || '').trim()); }
function normalizarCorreo_(c) { return String(c || '').trim().toLowerCase(); }
function telefonoValido_(t) { return String(t || '').replace(/\D/g, '').length >= 8; }

/** Cédula de identidad: 9 dígitos (X-XXXX-XXXX). DIMEX: 11 o 12 dígitos. Se guarda solo con números. */
function cedulaLimpia_(c) { return String(c == null ? '' : c).replace(/[\s.\-]/g, ''); }
function cedulaValida_(c) { return /^[1-9]\d{8}$/.test(c) || /^\d{11,12}$/.test(c); }

function hash_(clave, salt) {
  let texto = salt + '|' + clave;
  let bytes = null;
  for (let i = 0; i < ITER_HASH; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
    texto = Utilities.base64Encode(bytes) + salt;
  }
  return Utilities.base64Encode(bytes);
}

/**
 * Revisa que el nombre de usuario, la cédula y el correo no los tenga otra cuenta
 * (se puede ingresar con cualquiera de los tres, así que no se pueden repetir entre cuentas).
 */
function revisarIdentificadores_(usuarios, idPropio, d) {
  const otras = usuarios.filter(function (x) { return x.ID !== idPropio; });
  if (d.usuario && otras.some(function (x) { return x.Usuario === d.usuario; })) throw new Error('Ese nombre de usuario ya está en uso. Probá con otro.');
  if (d.cedula && otras.some(function (x) { return x.ID === d.cedula || x.Cedula === d.cedula; })) throw new Error('Ya existe una cuenta con esa cédula.');
  if (d.correo && otras.some(function (x) { return x.Correo === d.correo; })) throw new Error('Ya existe una cuenta con ese correo.');
}

function crearUsuario_(id, d, clave, rol, estado) {
  const salt = Utilities.getUuid();
  const u = {
    ID: id, Usuario: normalizarUsuario_(d.usuario), Cedula: cedulaLimpia_(d.cedula), Nombre: texto_(d.nombre, 120), Correo: normalizarCorreo_(d.correo),
    Avisos_correo: d.avisosCorreo === false ? 'No' : 'Sí', Telefono: texto_(d.telefono, 30),
    Rol: rol, Estado: estado, Hash: hash_(clave, salt), Salt: salt, Creado: ahora_(),
    Consentimiento: d.consentimiento ? ahora_() : ''
  };
  agregar_('Usuarios', u);
  return u;
}

function publico_(u) {
  return { id: u.ID, usuario: String(u.Usuario || ''), cedula: String(u.Cedula || ''), nombre: String(u.Nombre), correo: u.Correo,
           avisosCorreo: u.Avisos_correo !== 'No', telefono: String(u.Telefono || ''), rol: u.Rol,
           estado: u.Estado, creado: fechaHora_(u.Creado) };
}

/** Valida y limpia los datos del perfil de una mamá. Lanza un error con el primer problema. */
function limpiarPerfil_(p) {
  p = p || {};
  const hoy = hoyYmd_();
  const r = {
    cedula: cedulaLimpia_(p.cedula), usuario: normalizarUsuario_(p.usuario), avisosCorreo: p.avisosCorreo !== false,
    nombre: texto_(p.nombre, 120), telefono: texto_(p.telefono, 30),
    fechaNacimiento: String(p.fechaNacimiento || ''), universidad: texto_(p.universidad, 150),
    carrera: texto_(p.carrera, 150), beca: Number(p.beca), cuatrimestres: Number(p.cuatrimestres),
    hijos: Array.isArray(p.hijos) ? p.hijos : []
  };
  if (!cedulaValida_(r.cedula)) throw new Error('Revisá tu cédula: 9 dígitos (por ejemplo 1-0234-0567) o DIMEX de 11 o 12 dígitos.');
  if (!usuarioValido_(r.usuario)) throw new Error('El nombre de usuario debe tener de 3 a 30 caracteres, al menos una letra, y solo letras, números, punto, guion o guion bajo (sin espacios ni tildes).');
  if (r.nombre.length < 3) throw new Error('Escribí tu nombre completo.');
  if (!telefonoValido_(r.telefono)) throw new Error('El número de teléfono es obligatorio (al menos 8 dígitos).');
  if (!fechaValida_(r.fechaNacimiento) || r.fechaNacimiento >= hoy || edad_(r.fechaNacimiento, hoy) < 15 || edad_(r.fechaNacimiento, hoy) > 80) {
    throw new Error('Revisá tu fecha de nacimiento.');
  }
  if (r.universidad.length < 2) throw new Error('Escribí el nombre de tu universidad.');
  if (r.carrera.length < 2) throw new Error('Escribí la carrera que cursás.');
  if (p.beca === '' || p.beca == null || isNaN(r.beca) || r.beca < 0 || r.beca > 100) throw new Error('El porcentaje de beca debe estar entre 0 y 100 (0 si no tenés beca).');
  if (p.cuatrimestres === '' || p.cuatrimestres == null || !Number.isInteger(r.cuatrimestres) || r.cuatrimestres < 0 || r.cuatrimestres > 40) {
    throw new Error('Los cuatrimestres cursados deben ser un número entero entre 0 y 40.');
  }
  if (r.hijos.length > 15) throw new Error('Revisá la cantidad de hijos.');
  const vistas = [r.cedula];
  r.hijos = r.hijos.map(function (h, i) {
    h = h || {};
    const n = '#' + (i + 1);
    const hijo = { nombre: texto_(h.nombre, 120), cedula: cedulaLimpia_(h.cedula), fechaNacimiento: String(h.fechaNacimiento || ''),
                   genero: String(h.genero || '') };
    if (hijo.nombre.length < 2) throw new Error('Escribí el nombre del hijo o hija ' + n + '.');
    if (!cedulaValida_(hijo.cedula)) throw new Error('Revisá la cédula del hijo o hija ' + n + ' (9 dígitos o DIMEX de 11 o 12).');
    if (vistas.indexOf(hijo.cedula) !== -1) throw new Error('La cédula del hijo o hija ' + n + ' está repetida.');
    vistas.push(hijo.cedula);
    if (!fechaValida_(hijo.fechaNacimiento) || hijo.fechaNacimiento > hoy || edad_(hijo.fechaNacimiento, hoy) > 40) {
      throw new Error('Revisá la fecha de nacimiento del hijo o hija ' + n + '.');
    }
    if (GENEROS.indexOf(hijo.genero) === -1) throw new Error('Elegí el género del hijo o hija ' + n + '.');
    return hijo;
  });
  return r;
}

/** Cambia el identificador de una mamá en todas las hojas (por ejemplo, al registrar o corregir su cédula). */
function cambiarIdMama_(viejo, nuevo) {
  const cambios = [['Usuarios', 'ID'], ['Avisos', 'ID_Usuario'], ['Mamas', 'Cedula'], ['Hijos', 'Cedula_mama']]
    .concat(TABLAS_SOL.concat(['Mantenimiento']).map(function (k) { return [k, 'Cedula_mama']; }));
  cambios.forEach(function (par) {
    filas_(par[0]).forEach(function (r) {
      if (r[par[1]] === viejo) actualizarCelda_(par[0], r._fila, par[1], nuevo);
    });
  });
}

function guardarPerfil_(u, p, token) {
  let id = u.ID;
  revisarIdentificadores_(filas_('Usuarios'), u.ID, { usuario: p.usuario, cedula: p.cedula });
  if (p.cedula !== id) {
    cambiarIdMama_(id, p.cedula);
    id = p.cedula;
    if (token) CacheService.getScriptCache().put('s_' + token, JSON.stringify({ id: id }), SESION_SEG);
  }
  actualizarFila_('Usuarios', u._fila, { Nombre: p.nombre, Telefono: p.telefono, Usuario: p.usuario, Cedula: id,
                                         Avisos_correo: p.avisosCorreo ? 'Sí' : 'No' });
  const fila = {
    Cedula: id, Nombre: p.nombre, Correo: u.Correo, Telefono: p.telefono,
    Fecha_nacimiento: aFecha_(p.fechaNacimiento), Universidad: p.universidad, Carrera: p.carrera,
    Porcentaje_beca: p.beca, Cuatrimestres: p.cuatrimestres, Cantidad_hijos: p.hijos.length,
    Hijos: p.hijos.map(function (h) { return h.nombre + ' (' + h.genero + ')'; }).join('; '),
    Estado_cuenta: u.Estado, Actualizado: ahora_()
  };
  const existente = filas_('Mamas').find(function (r) { return r.Cedula === id; });
  if (existente) actualizarFila_('Mamas', existente._fila, fila);
  else { fila.Creado = ahora_(); agregar_('Mamas', fila); }
  quitarHijos_(id);
  p.hijos.forEach(function (h) {
    agregar_('Hijos', { Cedula_mama: id, Nombre_mama: p.nombre, Nombre_hijo: h.nombre, Cedula_hijo: h.cedula,
                        Fecha_nacimiento: aFecha_(h.fechaNacimiento), Genero: h.genero });
  });
  refrescarEdades_('Mamas');
  refrescarEdades_('Hijos');
  return id;
}

function quitarHijos_(idMama) {
  borrarFilas_('Hijos', filas_('Hijos').filter(function (r) { return r.Cedula_mama === idMama; }).map(function (r) { return r._fila; }));
}

function perfilDe_(idMama, filasMamas, filasHijos, usuario) {
  const m = (filasMamas || filas_('Mamas')).find(function (r) { return r.Cedula === idMama; });
  const hijos = (filasHijos || filas_('Hijos')).filter(function (r) { return r.Cedula_mama === idMama; })
    .map(function (r) {
      return { nombre: String(r.Nombre_hijo || ''), cedula: String(r.Cedula_hijo || ''),
               fechaNacimiento: ymd_(r.Fecha_nacimiento), genero: String(r.Genero || '') };
    });
  const cedula = cedulaValida_(idMama) ? idMama : '';
  const cuenta = { usuario: usuario ? String(usuario.Usuario || '') : '', avisosCorreo: usuario ? usuario.Avisos_correo !== 'No' : true };
  if (!m) return { completo: false, cedula: cedula, usuario: cuenta.usuario, avisosCorreo: cuenta.avisosCorreo, hijos: [], faltan: ['todos los datos del perfil'] };
  const p = {
    cedula: cedula, usuario: cuenta.usuario, avisosCorreo: cuenta.avisosCorreo, nombre: String(m.Nombre || ''), telefono: String(m.Telefono || ''), fechaNacimiento: ymd_(m.Fecha_nacimiento),
    universidad: String(m.Universidad || ''), carrera: String(m.Carrera || ''),
    beca: m.Porcentaje_beca === '' ? '' : Number(m.Porcentaje_beca),
    cuatrimestres: m.Cuatrimestres === '' ? '' : Number(m.Cuatrimestres),
    hijos: hijos, actualizado: fechaHora_(m.Actualizado)
  };
  const faltan = [];
  if (usuario && !usuarioValido_(p.usuario)) faltan.push('crear tu nombre de usuario');
  if (!p.cedula) faltan.push('tu cédula');
  if (!telefonoValido_(p.telefono)) faltan.push('tu teléfono');
  if (!p.fechaNacimiento) faltan.push('tu fecha de nacimiento');
  if (!p.universidad) faltan.push('tu universidad');
  if (!p.carrera) faltan.push('tu carrera');
  if (p.beca === '') faltan.push('el porcentaje de beca');
  if (p.cuatrimestres === '') faltan.push('los cuatrimestres cursados');
  hijos.forEach(function (h, i) {
    if (!cedulaValida_(h.cedula) || !h.fechaNacimiento || GENEROS.indexOf(h.genero) === -1) {
      faltan.push('la cédula, fecha de nacimiento o género de ' + (h.nombre || 'tu hijo o hija #' + (i + 1)));
    }
  });
  p.faltan = faltan;
  p.completo = faltan.length === 0;
  return p;
}

function sesion_(token, rol) {
  asegurarEstructura_();
  const raw = token ? CacheService.getScriptCache().get('s_' + token) : null;
  if (!raw) throw new Error('SESION: Tu sesión venció. Ingresá de nuevo.');
  const s = JSON.parse(raw);
  const u = filas_('Usuarios').find(function (x) { return x.ID === s.id; });
  if (!u || u.Estado !== 'Activa') throw new Error('SESION: Tu cuenta no está activa.');
  if (rol && (Array.isArray(rol) ? rol.indexOf(u.Rol) === -1 : u.Rol !== rol)) throw new Error('No tenés permiso para esta acción.');
  return u;
}

function registrar(d) {
  asegurarEstructura_();
  d = d || {};
  const p = limpiarPerfil_(d);
  if (!correoValido_(d.correo)) throw new Error('El correo no es válido.');
  if (String(d.clave || '').length < 8) throw new Error('La contraseña debe tener al menos 8 caracteres.');
  if (!d.consentimiento) throw new Error('Debés aceptar el uso de tus datos para crear la cuenta.');
  return conBloqueo_(function () {
    const correo = normalizarCorreo_(d.correo);
    const usuarios = filas_('Usuarios');
    revisarIdentificadores_(usuarios, null, { usuario: p.usuario, cedula: p.cedula, correo: correo });
    crearUsuario_(p.cedula, { cedula: p.cedula, nombre: p.nombre, correo: correo, telefono: p.telefono, usuario: p.usuario, avisosCorreo: p.avisosCorreo,
                              consentimiento: true }, d.clave, 'mama', 'Pendiente');
    const u = filas_('Usuarios').find(function (x) { return x.ID === p.cedula; });
    guardarPerfil_(u, p);
    avisarAdmins_('Nueva cuenta por activar: ' + p.nombre + ' (cédula ' + p.cedula + ', ' + correo + ', tel. ' + p.telefono + ').');
    return 'Cuenta creada. La administración debe activarla antes de que puedas ingresar.';
  });
}

function actualizarPerfil(token, d) {
  const u = sesion_(token, 'mama');
  const p = limpiarPerfil_(d);
  return conBloqueo_(function () {
    guardarPerfil_(u, p, token);
    return 'Perfil actualizado.';
  });
}

/** Mamás y administración pueden ingresar con su cédula, su nombre de usuario o su correo. */
function iniciarSesion(identificador, clave) {
  asegurarEstructura_();
  const ident = String(identificador || '').trim().toLowerCase();
  const ced = cedulaLimpia_(ident);
  const cache = CacheService.getScriptCache();
  const intentos = Number(cache.get('f_' + ident) || 0);
  if (intentos >= MAX_INTENTOS) throw new Error('Demasiados intentos fallidos. Esperá 15 minutos.');
  const u = ident && filas_('Usuarios').find(function (x) {
    return (x.Usuario && x.Usuario === ident) || x.Correo === ident ||
           (cedulaValida_(ced) && (x.ID === ced || x.Cedula === ced));
  });
  if (!u || hash_(String(clave || ''), u.Salt) !== u.Hash) {
    cache.put('f_' + ident, String(intentos + 1), BLOQUEO_SEG);
    throw new Error('Los datos de ingreso o la contraseña no son correctos.');
  }
  if (u.Estado === 'Pendiente') throw new Error('Tu cuenta todavía no ha sido activada por la administración.');
  if (u.Estado !== 'Activa') throw new Error('Tu cuenta está desactivada. Comunicate con la administración.');
  cache.remove('f_' + ident);
  const token = Utilities.getUuid();
  cache.put('s_' + token, JSON.stringify({ id: u.ID }), SESION_SEG);
  return { token: token, usuario: publico_(u) };
}

function cerrarSesion(token) {
  if (token) CacheService.getScriptCache().remove('s_' + token);
  return true;
}

function cambiarClave(token, actual, nueva) {
  const u = sesion_(token);
  if (hash_(String(actual || ''), u.Salt) !== u.Hash) throw new Error('La contraseña actual no es correcta.');
  if (String(nueva || '').length < 8) throw new Error('La nueva contraseña debe tener al menos 8 caracteres.');
  conBloqueo_(function () {
    const salt = Utilities.getUuid();
    actualizarFila_('Usuarios', u._fila, { Salt: salt, Hash: hash_(nueva, salt) });
  });
  return 'Contraseña actualizada.';
}

// ---------------------------------------------------------------------------
// Notificaciones y correos
// ---------------------------------------------------------------------------
function enviarCorreo_(para, asunto, cuerpo) {
  try {
    MailApp.sendEmail({ to: para, subject: asunto, body: cuerpo, name: 'Proyecto Bethany' });
  } catch (e) {
    Logger.log('No se pudo enviar el correo a %s: %s', para, e);
  }
}

function avisar_(usuario, mensaje, asunto) {
  agregar_('Avisos', { ID: 'A-' + Utilities.getUuid().slice(0, 8), Fecha: ahora_(), ID_Usuario: usuario.ID,
                       Mensaje: texto_(mensaje, 2000), Leido: 'No' });
  if (config_().correosMamas && usuario.Avisos_correo !== 'No') {
    enviarCorreo_(usuario.Correo, asunto || 'Proyecto Bethany: actualización de tu solicitud',
                  'Hola ' + usuario.Nombre + ',\n\n' + mensaje + '\n\nPodés ver el detalle en la aplicación de Proyecto Bethany.');
  }
}

/** Avisa por correo a la administración y al personal de mantenimiento activo. */
function avisarMantenimiento_(mensaje) {
  filas_('Usuarios').forEach(function (u) {
    if ((u.Rol === 'admin' || u.Rol === 'mant') && u.Estado === 'Activa') enviarCorreo_(u.Correo, 'Proyecto Bethany: solicitud de mantenimiento', mensaje);
  });
}

function avisarAdmins_(mensaje) {
  filas_('Usuarios').forEach(function (u) {
    if (u.Rol === 'admin' && u.Estado === 'Activa') enviarCorreo_(u.Correo, 'Proyecto Bethany: aviso del sistema', mensaje);
  });
}

function textoSalida_(s) {
  return 'del ' + s.salida + ' a las ' + s.horaSalida + ' al ' + s.regreso + ' a las ' + s.horaRegreso;
}

// ---------------------------------------------------------------------------
// Funciones de la mamá
// ---------------------------------------------------------------------------
function datosMama(token) {
  const u = sesion_(token, 'mama');
  const cfg = config_();
  const mias = filas_('Todas').filter(function (r) { return r.Cedula_mama === u.ID; }).map(mapaSolicitud_);
  const de = function (estados) { return mias.filter(function (s) { return estados.indexOf(s.estado) !== -1; }); };
  const pendientes = de(['Pendiente']);
  const aprobadas = de(['Aprobada']);
  const avisos = filas_('Avisos').filter(function (a) { return a.ID_Usuario === u.ID; })
    .map(function (a) { return { fecha: fechaHora_(a.Fecha), mensaje: String(a.Mensaje), leido: a.Leido === 'Sí' }; })
    .reverse().slice(0, 30);
  const hoy = hoyYmd_();
  return {
    usuario: publico_(u), perfil: perfilDe_(u.ID, null, null, u), generos: GENEROS,
    tokens: cfg.tokens, excepcion: cfg.excepcion, diasSalida: cfg.diasSalida, hoy: hoy,
    cuatri: semanaCuatri_(cfg, hoy),
    usos: tokensPorMes_(pendientes.concat(aprobadas)),
    pendientes: pendientes, aprobadas: aprobadas,
    rechazadas: de(['Rechazada']).reverse().slice(0, 10),
    otras: de(['Cancelada', 'Anulada']).reverse().slice(0, 10),
    avisos: avisos,
    mantenimiento: mantDeMama_(u.ID), desdeMant: DESDE_MANT
  };
}

function enviarSolicitud(token, d) {
  const u = sesion_(token, 'mama');
  d = d || {};
  const perfil = perfilDe_(u.ID, null, null, u);
  if (!perfil.completo) throw new Error('Completá tu perfil antes de enviar solicitudes.');
  const comentario = String(d.comentario || '').trim();
  if (comentario.length < 5) throw new Error('El comentario para la administración es obligatorio.');
  if (!horaValida_(d.horaSalida) || !horaValida_(d.horaRegreso)) throw new Error('Indicá la hora de salida y la hora de regreso.');
  return conBloqueo_(function () {
    const v = validarSalida_(u.ID, String(d.salida), String(d.regreso), null, { usarExcepcion: d.usarExcepcion === true });
    if (v.error) throw new Error(v.error);
    const exc = v.token === TOKEN_EXC;
    const fila = {
      ID: 'S-' + Utilities.getUuid().slice(0, 8), Enviada: ahora_(), Cedula_mama: u.ID, Nombre: u.Nombre,
      Correo: u.Correo, Telefono: u.Telefono,
      Salida: aFecha_(d.salida), Hora_salida: d.horaSalida, Regreso: aFecha_(d.regreso), Hora_regreso: d.horaRegreso,
      Dias: v.dias, Token: v.token, Semana: numeroSemana_(config_(), String(d.salida)),
      Comentario: texto_(comentario, 500), Estado: 'Pendiente'
    };
    agregar_('Todas', fila);
    agregar_('SolPendientes', fila);
    avisarAdmins_((exc ? '[TOKEN DE EXCEPCIÓN] ' : '') + 'Nueva solicitud de salida de ' + u.Nombre + ' ' +
                  textoSalida_({ salida: d.salida, horaSalida: d.horaSalida, regreso: d.regreso, horaRegreso: d.horaRegreso }) +
                  ' (' + v.dias + ' días).' +
                  (exc ? '\nEsta mamá ya usó todos sus tokens del mes y está usando su TOKEN DE EXCEPCIÓN.' : '') +
                  '\nComentario: ' + comentario);
    return exc ? 'Solicitud enviada con tu token de excepción.' : 'Solicitud enviada a la administración.';
  });
}

function cancelarSolicitud(token, id) {
  const u = sesion_(token, 'mama');
  return conBloqueo_(function () {
    const r = filas_('Todas').find(function (x) { return x.ID === id && x.Cedula_mama === u.ID && x.Estado === 'Pendiente'; });
    if (!r) throw new Error('La solicitud ya no está pendiente.');
    actualizarFila_('Todas', r._fila, { Estado: 'Cancelada', Revisada_por: 'Cancelada por la mamá', Fecha_revision: ahora_() });
    quitarDe_('SolPendientes', id);
    return 'Solicitud cancelada.';
  });
}

function marcarAvisosLeidos(token) {
  const u = sesion_(token);
  conBloqueo_(function () {
    filas_('Avisos').forEach(function (a) {
      if (a.ID_Usuario === u.ID && a.Leido !== 'Sí') actualizarCelda_('Avisos', a._fila, 'Leido', 'Sí');
    });
  });
  return true;
}

// ---------------------------------------------------------------------------
// Funciones de administración
// ---------------------------------------------------------------------------
function datosAdmin(token) {
  const u = sesion_(token, 'admin');
  const cfg = config_();
  const todas = filas_('Todas').map(mapaSolicitud_);
  const aprobadas = todas.filter(function (s) { return s.estado === 'Aprobada'; });
  const hoy = hoyYmd_();
  const conSemana = function (s) { if (s.semana === '') s.semana = numeroSemana_(cfg, s.salida); return s; };
  todas.forEach(conSemana);
  const pendientes = todas.filter(function (s) { return s.estado === 'Pendiente'; }).map(function (s) {
    const mes = s.salida.slice(0, 7);
    const uso = tokensPorMes_(aprobadas.filter(function (a) { return a.idMama === s.idMama; }))[mes] || { normal: 0, excepcion: 0 };
    s.saldo = { mes: mes, normal: uso.normal, excepcion: uso.excepcion };
    return s;
  });
  const rechazadas = todas.filter(function (s) { return s.estado === 'Rechazada'; }).reverse().slice(0, 100);
  const fm = filas_('Mamas'), fh = filas_('Hijos');
  const usuarios = filas_('Usuarios').map(function (x) {
    const p = publico_(x);
    if (x.Rol === 'mama') p.perfil = perfilDe_(x.ID, fm, fh, x);
    return p;
  });
  const idsUsuarios = usuarios.map(function (x) { return x.id; });
  const eliminadas = fm.filter(function (r) { return r.Estado_cuenta === 'Eliminada' && idsUsuarios.indexOf(r.Cedula) === -1; })
    .map(function (r) {
      return { id: r.Cedula, nombre: String(r.Nombre || ''), correo: String(r.Correo || ''), eliminada: fechaHora_(r.Actualizado),
               solicitudes: todas.filter(function (s) { return s.idMama === r.Cedula; }).length,
               hijos: fh.filter(function (h) { return h.Cedula_mama === r.Cedula; }).length };
    });
  const props = PropertiesService.getScriptProperties();
  const url = function (prop) { return 'https://docs.google.com/spreadsheets/d/' + props.getProperty(prop) + '/edit'; };
  return {
    usuario: publico_(u), tokens: cfg.tokens, excepcion: cfg.excepcion, diasSalida: cfg.diasSalida, hoy: hoy,
    correosMamas: cfg.correosMamas, eliminadas: eliminadas, cuatri: semanaCuatri_(cfg, hoy), semanas: SEMANAS_CUATRI,
    pendientes: pendientes, aprobadas: aprobadas, rechazadas: rechazadas, usuarios: usuarios,
    mantenimiento: filas_('Mantenimiento').map(function (r) { return mapaMant_(r, true); }).reverse(),
    estadosMant: ESTADOS_MANT, prioridadesMant: PRIORIDADES_MANT, desdeMant: DESDE_MANT,
    enlaces: { solicitudes: url('ID_SOLICITUDES'), aprobadas: url('ID_APROBADAS'), residentes: url('ID_RESIDENTES'),
               mantenimiento: url('ID_MANTENIMIENTO'), registroMant: url('ID_REG_MANT') }
  };
}

function decidir(token, id, decision, nota) {
  const admin = sesion_(token, 'admin');
  const notaTxt = texto_(nota, 500);
  return conBloqueo_(function () {
    const r = filas_('Todas').find(function (x) { return x.ID === id; });
    if (!r || r.Estado !== 'Pendiente') throw new Error('La solicitud ya fue atendida o cancelada.');
    const s = mapaSolicitud_(r);
    const mama = filas_('Usuarios').find(function (x) { return x.ID === s.idMama; });
    const revision = { Revisada_por: admin.Nombre, Fecha_revision: ahora_(), Nota_admin: notaTxt };
    if (decision === 'aprobar') {
      const v = validarSalida_(s.idMama, s.salida, s.regreso, s.id, { alAprobar: true, token: s.token });
      if (v.error) throw new Error('No se puede aprobar: ' + v.error);
      revision.Estado = 'Aprobada';
      actualizarFila_('Todas', r._fila, revision);
      quitarDe_('SolPendientes', id);
      agregar_('SolAprobadas', copiaSol_(r, revision));
      agregar_('Registro', copiaSol_(r, revision));
      if (mama) avisar_(mama, 'Tu salida ' + textoSalida_(s) + ' (' + s.dias + ' días) fue APROBADA' +
                         (s.token === TOKEN_EXC ? ' con tu token de excepción.' : '.') +
                         (notaTxt ? '\nNota de la administración: ' + notaTxt : ''));
      return 'Solicitud aprobada.';
    }
    if (decision === 'rechazar') {
      revision.Estado = 'Rechazada';
      actualizarFila_('Todas', r._fila, revision);
      quitarDe_('SolPendientes', id);
      agregar_('SolRechazadas', copiaSol_(r, revision));
      if (mama) avisar_(mama, 'Tu salida ' + textoSalida_(s) + ' fue RECHAZADA.' +
                         (notaTxt ? '\nMotivo: ' + notaTxt : '') +
                         '\nEl token se te devolvió: podés enviar una nueva solicitud.');
      return 'Solicitud rechazada.';
    }
    throw new Error('Decisión no válida.');
  });
}

function anularAprobada(token, id, nota) {
  const admin = sesion_(token, 'admin');
  const notaTxt = texto_(nota, 500);
  return conBloqueo_(function () {
    const r = filas_('Todas').find(function (x) { return x.ID === id && x.Estado === 'Aprobada'; });
    if (!r) throw new Error('La salida ya no está aprobada.');
    const s = mapaSolicitud_(r);
    actualizarFila_('Todas', r._fila, { Estado: 'Anulada', Revisada_por: admin.Nombre, Fecha_revision: ahora_(),
                                         Nota_admin: 'Anulada: ' + notaTxt });
    quitarDe_('SolAprobadas', id);
    quitarDe_('Registro', id);
    const mama = filas_('Usuarios').find(function (x) { return x.ID === s.idMama; });
    if (mama) avisar_(mama, 'Tu salida aprobada ' + textoSalida_(s) + ' fue ANULADA y el token se te devolvió.' +
                       (notaTxt ? '\nMotivo: ' + notaTxt : ''));
    return 'Salida anulada.';
  });
}

function cambiarEstadoUsuario(token, idUsuario, estado) {
  const admin = sesion_(token, 'admin');
  if (['Activa', 'Inactiva'].indexOf(estado) === -1) throw new Error('Estado no válido.');
  if (idUsuario === admin.ID) throw new Error('No podés cambiar el estado de tu propia cuenta.');
  return conBloqueo_(function () {
    const u = filas_('Usuarios').find(function (x) { return x.ID === idUsuario; });
    if (!u) throw new Error('La cuenta no existe.');
    actualizarCelda_('Usuarios', u._fila, 'Estado', estado);
    const m = filas_('Mamas').find(function (x) { return x.Cedula === idUsuario; });
    if (m) actualizarCelda_('Mamas', m._fila, 'Estado_cuenta', estado);
    if (estado === 'Activa' && u.Estado === 'Pendiente') {
      enviarCorreo_(u.Correo, 'Proyecto Bethany: cuenta activada', 'Hola ' + u.Nombre + ', tu cuenta ya está activa y podés ingresar al sistema de salidas.');
    }
    return 'Cuenta ' + (estado === 'Activa' ? 'activada.' : 'desactivada.');
  });
}

function restablecerClave(token, idUsuario, nueva) {
  sesion_(token, 'admin');
  if (String(nueva || '').length < 8) throw new Error('La contraseña temporal debe tener al menos 8 caracteres.');
  return conBloqueo_(function () {
    const u = filas_('Usuarios').find(function (x) { return x.ID === idUsuario; });
    if (!u) throw new Error('La cuenta no existe.');
    const salt = Utilities.getUuid();
    actualizarFila_('Usuarios', u._fila, { Salt: salt, Hash: hash_(nueva, salt) });
    return 'Contraseña temporal asignada. Comunicásela a ' + u.Nombre + ' por un medio seguro.';
  });
}

/** Crea una cuenta de administración o de personal de mantenimiento (solo la administración puede hacerlo). */
/** Valida los datos de una cuenta de administración o de personal; devuelve usuario, cédula y correo normalizados. */
function validarCuenta_(d) {
  if (String(d.nombre || '').trim().length < 3) throw new Error('Escribí el nombre completo.');
  if (!correoValido_(d.correo)) throw new Error('El correo no es válido.');
  if (String(d.clave || '').length < 8) throw new Error('La contraseña debe tener al menos 8 caracteres.');
  const usuario = normalizarUsuario_(d.usuario);
  if (!usuarioValido_(usuario)) throw new Error('El nombre de usuario debe tener de 3 a 30 caracteres, al menos una letra, y solo letras, números, punto, guion o guion bajo.');
  const cedula = cedulaLimpia_(d.cedula);
  if (cedula && !cedulaValida_(cedula)) throw new Error('Revisá la cédula: 9 dígitos (por ejemplo 1-0234-0567) o DIMEX de 11 o 12 dígitos.');
  if (d.telefono && !telefonoValido_(d.telefono)) throw new Error('Revisá el teléfono (al menos 8 dígitos).');
  return { usuario: usuario, cedula: cedula, correo: normalizarCorreo_(d.correo) };
}

function crearCuenta_(token, d, rol) {
  sesion_(token, 'admin');
  d = d || {};
  if (String(d.nombre || '').trim().length < 3) throw new Error('Escribí el nombre completo.');
  if (!correoValido_(d.correo)) throw new Error('El correo no es válido.');
  if (String(d.clave || '').length < 8) throw new Error('La contraseña debe tener al menos 8 caracteres.');
  const usuario = normalizarUsuario_(d.usuario);
  if (!usuarioValido_(usuario)) throw new Error('El nombre de usuario debe tener de 3 a 30 caracteres, al menos una letra, y solo letras, números, punto, guion o guion bajo.');
  const cedula = cedulaLimpia_(d.cedula);
  if (cedula && !cedulaValida_(cedula)) throw new Error('Revisá la cédula: 9 dígitos (por ejemplo 1-0234-0567) o DIMEX de 11 o 12 dígitos.');
  if (d.telefono && !telefonoValido_(d.telefono)) throw new Error('Revisá el teléfono (al menos 8 dígitos).');
  return conBloqueo_(function () {
    const correo = normalizarCorreo_(d.correo);
    revisarIdentificadores_(filas_('Usuarios'), null, { usuario: usuario, cedula: cedula, correo: correo });
    crearUsuario_((rol === 'mant' ? 'T-' : 'U-') + Utilities.getUuid().slice(0, 8),
                  { nombre: d.nombre, correo: correo, telefono: d.telefono || '', usuario: usuario, cedula: cedula, consentimiento: true },
                  d.clave, rol, 'Activa');
    return rol === 'mant' ? 'Cuenta de personal de mantenimiento creada.' : 'Cuenta de administración creada.';
  });
}

function crearAdmin(token, d) { return crearCuenta_(token, d, 'admin'); }
function crearMantenimiento(token, d) { return crearCuenta_(token, d, 'mant'); }

/** Elimina una cuenta de personal de mantenimiento (las boletas que atendió no se tocan). */
function eliminarCuentaPersonal(token, id, confirmacion) {
  sesion_(token, 'admin');
  if (confirmacion !== 'ELIMINAR') throw new Error('Para confirmar escribí ELIMINAR en mayúsculas.');
  return conBloqueo_(function () {
    const u = filas_('Usuarios').find(function (x) { return x.ID === id; });
    if (!u) throw new Error('La cuenta ya no existe.');
    if (u.Rol !== 'mant') throw new Error('Solo se pueden eliminar aquí cuentas de personal de mantenimiento.');
    borrarFilas_('Usuarios', [u._fila]);
    return 'Cuenta de ' + u.Nombre + ' eliminada.';
  });
}

function guardarConfig(token, tokensMes, tokensExc, diasSalida, correosMamas) {
  sesion_(token, 'admin');
  const t = Math.floor(Number(tokensMes)), e = Math.floor(Number(tokensExc));
  if (tokensMes === '' || !(t >= 0 && t <= 31)) throw new Error('Los tokens por mes deben estar entre 0 y 31.');
  if (tokensExc === '' || !(e >= 0 && e <= 31)) throw new Error('Los tokens de excepción deben estar entre 0 y 31.');
  const dias = (diasSalida || []).map(Number).filter(function (n) { return n >= 0 && n <= 6; });
  if (!dias.length) throw new Error('Elegí al menos un día de salida.');
  return conBloqueo_(function () {
    escribirConfig_('TOKENS_MES', t);
    escribirConfig_('TOKENS_EXCEPCION', e);
    escribirConfig_('DIAS_SALIDA', dias.map(function (n) { return NOMBRES_DIAS[n]; }).join(','));
    if (correosMamas !== undefined) escribirConfig_('CORREOS_MAMAS', correosMamas ? 'Sí' : 'No');
    return 'Configuración guardada.';
  });
}

/** Solo la administración define cuándo empieza la semana 1 y cuándo termina el cuatrimestre (semana 15). */
function guardarCuatrimestre(token, inicio, fin) {
  sesion_(token, 'admin');
  inicio = String(inicio || ''); fin = String(fin || '');
  if (!inicio && !fin) {
    return conBloqueo_(function () { escribirConfig_('CUATRI_INICIO', ''); escribirConfig_('CUATRI_FIN', ''); return 'Fechas del cuatrimestre borradas.'; });
  }
  if (!fechaValida_(inicio)) throw new Error('Indicá la fecha de inicio de la semana 1.');
  if (!fechaValida_(fin)) throw new Error('Indicá la fecha de fin del cuatrimestre (semana ' + SEMANAS_CUATRI + ').');
  const dias = (utc_(fin) - utc_(inicio)) / 86400000;
  if (dias < 7 * (SEMANAS_CUATRI - 1)) throw new Error('Entre el inicio y el fin debe haber al menos ' + (SEMANAS_CUATRI - 1) + ' semanas para llegar a la semana ' + SEMANAS_CUATRI + '.');
  if (dias > 7 * (SEMANAS_CUATRI + 4)) throw new Error('Revisá las fechas: el cuatrimestre parece demasiado largo.');
  return conBloqueo_(function () {
    escribirConfig_('CUATRI_INICIO', aFecha_(inicio));
    escribirConfig_('CUATRI_FIN', aFecha_(fin));
    const s = semanaCuatri_(config_(), hoyYmd_());
    return 'Cuatrimestre guardado.' + (s.estado === 'en_curso' ? ' Hoy es la semana ' + s.semana + '.' : '');
  });
}

/** Cada cuenta de administración puede editar su nombre, usuario, correo y cédula (para ingresar con cualquiera). */
function actualizarCuentaAdmin(token, d) {
  const u = sesion_(token, ['admin', 'mant']);
  d = d || {};
  const nombre = texto_(d.nombre, 120);
  const usuario = normalizarUsuario_(d.usuario);
  const correo = normalizarCorreo_(d.correo);
  const cedula = cedulaLimpia_(d.cedula);
  if (nombre.length < 3) throw new Error('Escribí el nombre completo.');
  if (!usuarioValido_(usuario)) throw new Error('El nombre de usuario debe tener de 3 a 30 caracteres, al menos una letra, y solo letras, números, punto, guion o guion bajo.');
  if (!correoValido_(correo)) throw new Error('El correo no es válido.');
  if (cedula && !cedulaValida_(cedula)) throw new Error('Revisá la cédula: 9 dígitos (por ejemplo 1-0234-0567) o DIMEX de 11 o 12 dígitos.');
  return conBloqueo_(function () {
    revisarIdentificadores_(filas_('Usuarios'), u.ID, { usuario: usuario, cedula: cedula, correo: correo });
    actualizarFila_('Usuarios', u._fila, { Nombre: nombre, Usuario: usuario, Correo: correo, Cedula: cedula });
    return { mensaje: 'Cuenta actualizada. Podés ingresar con tu usuario, tu correo' + (cedula ? ' o tu cédula.' : '.'),
             usuario: publico_(Object.assign({}, u, { Nombre: nombre, Usuario: usuario, Correo: correo, Cedula: cedula })) };
  });
}

/** Elimina la cuenta de una mamá (ya no puede ingresar). Sus datos quedan en los Excel hasta que se borren aparte. */
function eliminarCuentaMama(token, id, confirmacion) {
  const admin = sesion_(token, 'admin');
  if (confirmacion !== 'ELIMINAR') throw new Error('Para confirmar escribí ELIMINAR en mayúsculas.');
  return conBloqueo_(function () {
    const u = filas_('Usuarios').find(function (x) { return x.ID === id; });
    if (!u) throw new Error('La cuenta ya no existe.');
    if (u.Rol !== 'mama') throw new Error('Solo se pueden eliminar cuentas de mamás.');
    filas_('Todas').forEach(function (r) {
      if (r.Cedula_mama === id && r.Estado === 'Pendiente') {
        actualizarFila_('Todas', r._fila, { Estado: 'Cancelada', Revisada_por: 'Cuenta eliminada por ' + admin.Nombre, Fecha_revision: ahora_() });
      }
    });
    borrarFilas_('SolPendientes', filas_('SolPendientes').filter(function (r) { return r.Cedula_mama === id; }).map(function (r) { return r._fila; }));
    borrarFilas_('Avisos', filas_('Avisos').filter(function (r) { return r.ID_Usuario === id; }).map(function (r) { return r._fila; }));
    const m = filas_('Mamas').find(function (r) { return r.Cedula === id; });
    if (m) actualizarFila_('Mamas', m._fila, { Estado_cuenta: 'Eliminada', Actualizado: ahora_() });
    else agregar_('Mamas', { Cedula: id, Nombre: u.Nombre, Correo: u.Correo, Telefono: u.Telefono, Estado_cuenta: 'Eliminada', Actualizado: ahora_() });
    borrarFilas_('Usuarios', [u._fila]);
    return { mensaje: 'Cuenta de ' + u.Nombre + ' eliminada.', id: id, nombre: String(u.Nombre) };
  });
}

/** Borra definitivamente de los Excel todos los datos de una mamá cuya cuenta ya fue eliminada. */
function borrarDatosMama(token, id, confirmacion) {
  const admin = sesion_(token, 'admin');
  if (confirmacion !== 'ELIMINAR') throw new Error('Para confirmar escribí ELIMINAR en mayúsculas.');
  return conBloqueo_(function () {
    if (filas_('Usuarios').some(function (x) { return x.ID === id; })) throw new Error('Primero hay que eliminar la cuenta de la mamá.');
    const idsMant = filas_('Mantenimiento').filter(function (r) { return r.Cedula_mama === id; }).map(function (r) { return r.ID; });
    sincronizarRegistroMant_();   // el registro permanente de mantenimiento conserva sus boletas
    borrarFilas_('Firmas', filas_('Firmas').filter(function (r) { return idsMant.indexOf(r.ID) !== -1; }).map(function (r) { return r._fila; }));
    const objetivos = TABLAS_SOL.concat(['Mantenimiento']).map(function (k) { return [k, 'Cedula_mama']; })
      .concat([['Mamas', 'Cedula'], ['Hijos', 'Cedula_mama'], ['Avisos', 'ID_Usuario']]);
    let total = 0;
    objetivos.forEach(function (par) {
      const filas = filas_(par[0]).filter(function (r) { return r[par[1]] === id; }).map(function (r) { return r._fila; });
      total += filas.length;
      borrarFilas_(par[0], filas);
    });
    if (!total) throw new Error('No hay datos guardados de esa mamá.');
    if (idsMant.length) sincronizarRegistroMant_({ ids: idsMant, por: admin.Nombre + ' (datos de la mamá borrados)', fecha: ahora_() });
    refrescarEdades_('Mamas');
    refrescarEdades_('Hijos');
    return 'Datos borrados de los Excel (' + total + ' filas).';
  });
}

/** Envía un correo de prueba a la cuenta de administración que lo pide. */
function probarCorreo(token) {
  const admin = sesion_(token, 'admin');
  MailApp.sendEmail({ to: admin.Correo, subject: 'Proyecto Bethany: correo de prueba', name: 'Proyecto Bethany',
                      body: 'Este es un correo de prueba del sistema de salidas. Si lo recibiste, los avisos por correo funcionan.' });
  return 'Correo de prueba enviado a ' + admin.Correo + '. Correos disponibles hoy: ' + MailApp.getRemainingDailyQuota() + '.';
}

/** Devuelve un libro en formato .xlsx (base64) para descargarlo desde la aplicación. */
function descargarExcel(token, cual) {
  sesion_(token, 'admin');
  const prop = { solicitudes: 'ID_SOLICITUDES', aprobadas: 'ID_APROBADAS', residentes: 'ID_RESIDENTES', mantenimiento: 'ID_MANTENIMIENTO',
                 registroMant: 'ID_REG_MANT' }[cual];
  if (!prop) throw new Error('Libro no válido.');
  const id = PropertiesService.getScriptProperties().getProperty(prop);
  const nombre = DriveApp.getFileById(id).getName();
  const resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('No se pudo generar el archivo de Excel.');
  return { nombre: nombre + ' ' + hoyYmd_() + '.xlsx', base64: Utilities.base64Encode(resp.getContent()) };
}

// ---------------------------------------------------------------------------
// Mantenimiento y reparaciones
// Cada mamá llena la boleta (secciones 1 a 3) desde la app. La administración la ve completa en la app,
// llena la sección 4 (uso exclusivo del equipo de mantenimiento) y, si quiere, descarga la boleta en .docx.
// ---------------------------------------------------------------------------
function siNo_(v) { return v === 'Sí' || v === true ? 'Sí' : v === 'No' || v === false ? 'No' : ''; }
function fechaOpc_(v) { const f = ymd_(v); return fechaValida_(f) ? f : ''; }

function mapaMant_(r, completo, personal) {
  const o = {
    id: r.ID, enviada: fechaHora_(r.Enviada), idMama: r.Cedula_mama, cedula: r.Cedula_mama, nombre: String(r.Nombre || ''),
    telefono: String(r.Telefono || ''), correo: String(r.Correo || ''), apartamento: String(r.Apartamento || ''),
    descripcion: String(r.Descripcion || ''), fechaAproximada: fechaOpc_(r.Fecha_aproximada), desdeCuando: String(r.Desde_cuando || ''),
    riesgo: siNo_(r.Riesgo), riesgoDetalle: String(r.Riesgo_detalle || ''), inmediata: siNo_(r.Atencion_inmediata),
    infoAdicional: String(r.Info_adicional || ''), estado: ESTADOS_MANT.indexOf(r.Estado) === -1 ? 'Pendiente' : r.Estado,
    fechaAtencion: fechaOpc_(r.Fecha_atencion), fechaFinalizacion: fechaOpc_(r.Fecha_finalizacion)
  };
  if (completo) {
    Object.assign(o, {
      prioridad: PRIORIDADES_MANT.indexOf(r.Prioridad) === -1 ? '' : r.Prioridad, fechaRecepcion: fechaOpc_(r.Fecha_recepcion),
      evaluacion: String(r.Evaluacion || ''), trabajo: String(r.Trabajo || ''), materiales: String(r.Materiales || ''),
      responsable: String(r.Responsable || ''), observaciones: String(r.Observaciones || ''),
      firmaResponsable: r.Firma_responsable === 'Sí', actualizadaPor: String(r.Actualizada_por || ''), actualizada: fechaHora_(r.Actualizada)
    });
  }
  if (personal) { o.cedula = ''; o.idMama = ''; o.correo = ''; }   // el personal de mantenimiento solo necesita nombre, teléfono y apartamento
  return o;
}

/** La mamá ve sus boletas, su estado y las fechas de atención (no ve la evaluación interna de la sección 4). */
function mantDeMama_(idMama) {
  return filas_('Mantenimiento').filter(function (r) { return r.Cedula_mama === idMama; })
    .map(function (r) { return mapaMant_(r, false); }).reverse();
}

function firmaValida_(f) {
  return typeof f === 'string' && /^data:image\/png;base64,[A-Za-z0-9+\/=]+$/.test(f) && f.length <= MAX_FIRMA && f.length > 200;
}

function guardarFirma_(id, tipo, imagen) {
  const previa = filas_('Firmas').find(function (r) { return r.ID === id && r.Tipo === tipo; });
  if (previa) actualizarFila_('Firmas', previa._fila, { Imagen: imagen, Fecha: ahora_() });
  else agregar_('Firmas', { ID: id, Tipo: tipo, Imagen: imagen, Fecha: ahora_() });
}

function firmasDe_(id) {
  const res = { solicitante: '', responsable: '' };
  filas_('Firmas').forEach(function (r) {
    if (r.ID !== id) return;
    if (r.Tipo === 'Solicitante') res.solicitante = String(r.Imagen || '');
    if (r.Tipo === 'Responsable') res.responsable = String(r.Imagen || '');
  });
  return res;
}

/** La mamá envía la boleta de mantenimiento (secciones 1, 2 y 3, con su firma). */
function enviarMantenimiento(token, d) {
  const u = sesion_(token, 'mama');
  d = d || {};
  const perfil = perfilDe_(u.ID, null, null, u);
  if (!perfil.completo) throw new Error('Completá tu perfil antes de enviar solicitudes.');
  const apto = texto_(d.apartamento, 20);
  const desc = texto_(d.descripcion, 2000);
  const fechaAprox = String(d.fechaAproximada || '');
  if (!apto) throw new Error('Indicá el número de apartamento.');
  if (desc.length < 10) throw new Error('Describí el problema con el mayor detalle posible.');
  if (fechaAprox && (!fechaValida_(fechaAprox) || fechaAprox > hoyYmd_())) throw new Error('Revisá la fecha aproximada: no puede ser futura.');
  if (DESDE_MANT.indexOf(d.desdeCuando) === -1) throw new Error('Indicá desde cuándo se presenta el problema.');
  const riesgo = siNo_(d.riesgo), inmediata = siNo_(d.inmediata);
  if (!riesgo) throw new Error('Indicá si el problema está causando algún riesgo o afectación importante.');
  const riesgoDet = texto_(d.riesgoDetalle, 1000);
  if (riesgo === 'Sí' && riesgoDet.length < 3) throw new Error('Explicá brevemente el riesgo o la afectación.');
  if (!inmediata) throw new Error('Indicá si el problema requiere atención inmediata.');
  if (!firmaValida_(d.firma)) throw new Error('Falta tu firma. Firmá en el recuadro con el dedo o el mouse.');
  return conBloqueo_(function () {
    const id = 'M-' + Utilities.getUuid().slice(0, 8);
    const enviada = ahora_();
    agregar_('Mantenimiento', {
      ID: id, Enviada: enviada, Cedula_mama: u.ID, Nombre: u.Nombre, Telefono: u.Telefono, Correo: u.Correo, Apartamento: apto,
      Descripcion: desc, Fecha_aproximada: fechaAprox ? aFecha_(fechaAprox) : '', Desde_cuando: d.desdeCuando,
      Riesgo: riesgo, Riesgo_detalle: riesgo === 'Sí' ? riesgoDet : '', Atencion_inmediata: inmediata,
      Info_adicional: texto_(d.infoAdicional, 2000), Firma: 'Firmada en la app', Estado: 'Pendiente'
    });
    guardarFirma_(id, 'Solicitante', d.firma);
    const urgente = inmediata === 'Sí' || riesgo === 'Sí';
    sincronizarRegistroMant_();
    avisarMantenimiento_((urgente ? '[URGENTE] ' : '') + 'Nueva solicitud de mantenimiento de ' + u.Nombre + ', apartamento ' + apto +
                  ', enviada el ' + enviada + '.\n\nProblema: ' + desc +
                  (riesgo === 'Sí' ? '\nRiesgo o afectación: ' + riesgoDet : '') +
                  '\nRequiere atención inmediata: ' + inmediata +
                  '\n\nPodés ver la boleta completa y descargarla en .docx en la sección Mantenimiento de la app.');
    return 'Solicitud de mantenimiento enviada a la administración.';
  });
}

/** Sección 4: uso exclusivo del equipo de mantenimiento (administración y personal de mantenimiento). */
function guardarMantenimientoAdmin(token, id, d) {
  const admin = sesion_(token, ['admin', 'mant']);
  d = d || {};
  if (ESTADOS_MANT.indexOf(d.estado) === -1) throw new Error('Elegí el estado de la solicitud.');
  if (d.prioridad && PRIORIDADES_MANT.indexOf(d.prioridad) === -1) throw new Error('Prioridad no válida.');
  const fechas = { Fecha_recepcion: d.fechaRecepcion, Fecha_atencion: d.fechaAtencion, Fecha_finalizacion: d.fechaFinalizacion };
  Object.keys(fechas).forEach(function (k) {
    const f = String(fechas[k] || '');
    if (f && !fechaValida_(f)) throw new Error('Revisá las fechas de la sección 4.');
    fechas[k] = f;
  });
  if (d.firmaResponsable && !firmaValida_(d.firmaResponsable)) throw new Error('No se pudo leer la firma del responsable. Firmá de nuevo.');
  return conBloqueo_(function () {
    const r = filas_('Mantenimiento').find(function (x) { return x.ID === id; });
    if (!r) throw new Error('La solicitud de mantenimiento ya no existe.');
    const hoy = hoyYmd_();
    if (!fechas.Fecha_recepcion) fechas.Fecha_recepcion = fechaOpc_(r.Fecha_recepcion) || hoy;
    if (d.estado === 'Finalizada' && !fechas.Fecha_finalizacion) fechas.Fecha_finalizacion = hoy;
    const anterior = ESTADOS_MANT.indexOf(r.Estado) === -1 ? 'Pendiente' : r.Estado;
    const cambios = {
      Estado: d.estado, Prioridad: d.prioridad || '', Evaluacion: texto_(d.evaluacion, 2000), Trabajo: texto_(d.trabajo, 2000),
      Materiales: texto_(d.materiales, 2000), Responsable: texto_(d.responsable, 120), Observaciones: texto_(d.observaciones, 2000),
      Actualizada_por: admin.Nombre, Actualizada: ahora_()
    };
    Object.keys(fechas).forEach(function (k) { cambios[k] = fechas[k] ? aFecha_(fechas[k]) : ''; });
    if (d.firmaResponsable) { guardarFirma_(id, 'Responsable', d.firmaResponsable); cambios.Firma_responsable = 'Sí'; }
    if (d.borrarFirmaResponsable && !d.firmaResponsable) {
      borrarFilas_('Firmas', filas_('Firmas').filter(function (x) { return x.ID === id && x.Tipo === 'Responsable'; }).map(function (x) { return x._fila; }));
      cambios.Firma_responsable = '';
    }
    actualizarFila_('Mantenimiento', r._fila, cambios);
    if (d.estado !== anterior) {
      const mama = filas_('Usuarios').find(function (x) { return x.ID === r.Cedula_mama; });
      if (mama) {
        avisar_(mama, 'Tu solicitud de mantenimiento del apartamento ' + r.Apartamento + ' (enviada el ' + fechaHora_(r.Enviada) +
                ') ahora está: ' + d.estado.toUpperCase() + '.' +
                (fechas.Fecha_atencion ? '\nFecha de atención: ' + fechaCorta_(fechas.Fecha_atencion) + '.' : '') +
                (d.estado === 'Finalizada' ? '\nFecha de finalización: ' + fechaCorta_(fechas.Fecha_finalizacion) + '.' : ''),
                'Proyecto Bethany: actualización de tu solicitud de mantenimiento');
      }
    }
    sincronizarRegistroMant_();
    return 'Sección 4 guardada.' + (d.estado !== anterior ? ' Se avisó a la mamá del nuevo estado.' : '');
  });
}

/** Datos del perfil de personal de mantenimiento: solo las boletas (sin cuatrimestre, salidas ni perfiles). */
function datosMant(token) {
  const u = sesion_(token, 'mant');
  return {
    usuario: publico_(u), hoy: hoyYmd_(),
    mantenimiento: filas_('Mantenimiento').map(function (r) { return mapaMant_(r, true, true); }).reverse(),
    estadosMant: ESTADOS_MANT, prioridadesMant: PRIORIDADES_MANT, desdeMant: DESDE_MANT
  };
}

/** Devuelve los datos que corresponden al perfil de quien inició sesión. */
function datosSesion(token) {
  const u = sesion_(token);
  if (u.Rol === 'admin') return datosAdmin(token);
  if (u.Rol === 'mant') return datosMant(token);
  return datosMama(token);
}

/**
 * La administración borra de la app una boleta FINALIZADA. Su registro queda para siempre en el libro
 * «Bethany - Registro de mantenimiento», marcado como borrada de la app.
 */
function borrarBoleta(token, id, confirmacion) {
  const admin = sesion_(token, 'admin');
  const conf = String(confirmacion || '').trim().toUpperCase().replace('Í', 'I');
  if (conf !== 'SI') throw new Error('Para confirmar escribí SÍ.');
  return conBloqueo_(function () {
    sincronizarRegistroMant_();        // primero se asegura de que la boleta esté completa en el registro permanente
    const r = filas_('Mantenimiento').find(function (x) { return x.ID === id; });
    if (!r) throw new Error('La boleta ya no existe en la app.');
    if (r.Estado !== 'Finalizada') throw new Error('Solo se pueden borrar boletas finalizadas (con la sección 4 completa).');
    borrarFilas_('Firmas', filas_('Firmas').filter(function (x) { return x.ID === id; }).map(function (x) { return x._fila; }));
    borrarFilas_('Mantenimiento', [r._fila]);
    sincronizarRegistroMant_({ ids: [id], por: admin.Nombre, fecha: ahora_() });
    return 'Boleta borrada de la app. Su registro se conserva en el Excel «Registro de mantenimiento».';
  });
}

/**
 * Registro permanente de boletas. Copia el estado actual de cada boleta de la app, conserva las que ya no están
 * en la app (nunca borra filas) y deja la hoja «Todas (por estado)» ordenada por estado, más una hoja por estado.
 */
function sincronizarRegistroMant_(borrada) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ID_REG_MANT')) return;
  const orden = [], porId = {};
  filas_('RegTodas').forEach(function (r) {
    if (!r.ID) return;
    const o = {};
    COLS_REG_MANT.forEach(function (c) { o[c] = r[c] === undefined ? '' : r[c]; });
    if (!porId[r.ID]) orden.push(r.ID);
    porId[r.ID] = o;
  });
  filas_('Mantenimiento').forEach(function (r) {
    const o = porId[r.ID] || {};
    COLS_MANT.forEach(function (c) { o[c] = r[c] === undefined ? '' : r[c]; });
    o.En_la_app = 'Sí'; o.Borrada_por = ''; o.Fecha_borrado = '';
    if (!porId[r.ID]) orden.push(r.ID);
    porId[r.ID] = o;
  });
  const enApp = filas_('Mantenimiento').map(function (r) { return r.ID; });
  orden.forEach(function (id) {
    const o = porId[id];
    if (enApp.indexOf(id) !== -1) return;
    o.En_la_app = 'No (borrada)';
    if (borrada && borrada.ids.indexOf(id) !== -1) { o.Borrada_por = borrada.por; o.Fecha_borrado = borrada.fecha; }
  });
  const pos = function (e) { const i = ESTADOS_MANT.indexOf(e); return i === -1 ? 0 : i; };
  const clave = function (o) { return fechaHora_(o.Enviada); };
  const todas = orden.map(function (id) { return porId[id]; }).sort(function (a, b) {
    return pos(a.Estado) - pos(b.Estado) || (clave(a) < clave(b) ? 1 : clave(a) > clave(b) ? -1 : 0);
  });
  escribirTabla_('RegTodas', todas);
  Object.keys(HOJAS_REG_MANT).forEach(function (k) {
    escribirTabla_(k, todas.filter(function (o) { return (ESTADOS_MANT.indexOf(o.Estado) === -1 ? 'Pendiente' : o.Estado) === HOJAS_REG_MANT[k]; }));
  });
}

/** Reescribe una hoja completa (encabezado + filas) en una sola operación, con el color de cada estado. */
function escribirTabla_(k, objetos) {
  const sh = hoja_(k);
  const enc = TABLAS[k].cols;
  sh.clearContents();
  sh.clearFormats();
  const valores = [enc].concat(objetos.map(function (o) { return enc.map(function (c) { return valorCelda_(c, o[c]); }); }));
  sh.getRange(1, 1, valores.length, enc.length).setValues(valores);
  sh.getRange(1, 1, 1, enc.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (objetos.length) {
    sh.getRange(2, 1, objetos.length, enc.length).setBackgrounds(objetos.map(function (o) {
      const color = COLOR_MANT[o.Estado] || null;
      return enc.map(function () { return color; });
    }));
  }
  delete cache_[k];
}

/** Firmas de una boleta (se cargan solo al abrirla, para que la app no descargue todas las imágenes). */
function firmasMantenimiento(token, id) {
  sesion_(token, ['admin', 'mant']);
  return firmasDe_(id);
}

/** Genera la boleta en formato Word (.docx) para descargarla desde la app. */
function descargarMantenimientoDocx(token, id) {
  const u = sesion_(token, ['admin', 'mant']);
  const r = filas_('Mantenimiento').find(function (x) { return x.ID === id; });
  if (!r) throw new Error('La solicitud de mantenimiento ya no existe.');
  const m = mapaMant_(r, true, u.Rol === 'mant');
  const blob = docxMantenimiento_(m, firmasDe_(id));
  const nombre = 'Boleta de mantenimiento - Apto ' + m.apartamento.replace(/[\\\/:*?"<>|]/g, '-') + ' - ' + m.nombre + ' - ' + m.enviada.slice(0, 10) + '.docx';
  return { nombre: nombre, base64: Utilities.base64Encode(blob.getBytes()) };
}

function fechaCorta_(ymd) {
  const p = String(ymd || '').split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : '';
}

// ---------------------------------------------------------------------------
// Boleta en Word (.docx): se arma el archivo directamente (sin crear documentos en Drive ni pedir permisos nuevos)
// ---------------------------------------------------------------------------
const LOGO_DOCX = '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';

function xml_(s) {
  return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Un fragmento de texto con formato. o: b (negrita), i (cursiva), u (subrayado), color, tam (pt), fuente. */
function wr_(texto, o) {
  o = o || {};
  const pr = (o.fuente ? '<w:rFonts w:ascii="' + o.fuente + '" w:hAnsi="' + o.fuente + '" w:eastAsia="' + o.fuente + '" w:cs="' + o.fuente + '"/>' : '') +
    (o.b ? '<w:b/><w:bCs/>' : '') + (o.i ? '<w:i/><w:iCs/>' : '') + (o.color ? '<w:color w:val="' + o.color + '"/>' : '') +
    (o.tam ? '<w:sz w:val="' + o.tam * 2 + '"/><w:szCs w:val="' + o.tam * 2 + '"/>' : '') + (o.u ? '<w:u w:val="single"/>' : '');
  return String(texto == null ? '' : texto).split('\n').map(function (linea, i) {
    return '<w:r>' + (pr ? '<w:rPr>' + pr + '</w:rPr>' : '') + (i ? '<w:br/>' : '') + '<w:t xml:space="preserve">' + xml_(linea) + '</w:t></w:r>';
  }).join('');
}

/** Un párrafo. o: antes/despues (pt), alin, sangria (cm), linea (borde inferior), salto (página nueva antes), juntos. */
function wp_(runs, o) {
  o = o || {};
  const pr = (o.juntos ? '<w:keepNext/>' : '') + (o.salto ? '<w:pageBreakBefore/>' : '') +
    (o.linea ? '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="404040"/></w:pBdr>' : '') +
    '<w:spacing w:before="' + Math.round((o.antes || 0) * 20) + '" w:after="' + Math.round((o.despues === undefined ? 6 : o.despues) * 20) + '"/>' +
    (o.sangria ? '<w:ind w:left="' + Math.round(o.sangria * 567) + '"/>' : '') + (o.alin ? '<w:jc w:val="' + o.alin + '"/>' : '');
  return '<w:p><w:pPr>' + pr + '</w:pPr>' + (Array.isArray(runs) ? runs.join('') : runs || '') + '</w:p>';
}

function casilla_(marcada, texto) {
  return wr_(marcada ? '☒' : '☐', { fuente: 'Segoe UI Symbol' }) + wr_(' ' + texto);
}

function tamPng_(bytes) {
  const n = function (i) { return ((bytes[i] & 255) << 24 >>> 0) + ((bytes[i + 1] & 255) << 16) + ((bytes[i + 2] & 255) << 8) + (bytes[i + 3] & 255); };
  return { ancho: n(16) || 1, alto: n(20) || 1 };
}

/** Imagen en línea. anchoCm: ancho en el documento; el alto se calcula para no deformarla. */
function wimg_(rId, idImg, tam, anchoCm, nombre) {
  const cx = Math.round(anchoCm * 360000), cy = Math.round(cx * tam.alto / tam.ancho);
  return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
    '<wp:docPr id="' + idImg + '" name="' + nombre + '"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
    '<pic:nvPicPr><pic:cNvPr id="' + idImg + '" name="' + nombre + '.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="' + rId + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}

function docxMantenimiento_(m, firmas) {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const ROJO = '9B2335';
  const media = [];   // { rId, nombre, bytes }
  const imagen = function (dataUrl, nombre, anchoCm) {
    if (!dataUrl) return '';
    const bytes = Utilities.base64Decode(String(dataUrl).replace(/^data:image\/png;base64,/, ''));
    const rId = 'rIdImg' + (media.length + 1);
    media.push({ rId: rId, nombre: nombre + '.png', bytes: bytes });
    return wimg_(rId, media.length + 10, tamPng_(bytes), anchoCm, nombre);
  };
  const fecha = function (ymd) { return ymd ? wr_(fechaCorta_(ymd), { u: true }) : wr_('____ / ____ / ________'); };
  const valor = function (v) { return v ? wr_(v, { u: true }) : wr_('_______________________________'); };
  const bloque = function (v) { return wp_(wr_(v || ' '), { linea: true, despues: 10 }); };
  const titulo = function (t, o) { return wp_(wr_(t, Object.assign({ b: true, u: true }, o && o.run)), Object.assign({ antes: 12, despues: 8, juntos: true }, o && o.par)); };
  const preg = function (t, extra) { return wp_([wr_(t, { b: true }), extra ? wr_(' ' + extra) : ''], { antes: 10, despues: 4, juntos: true }); };
  const hora = m.enviada.slice(11, 16);
  const c = [];

  c.push(wp_(wr_('PROYECTO BETHANY', { b: true, tam: 14 }), { despues: 4 }));
  c.push(wp_(wr_('Boleta de Solicitud de Mantenimiento y Reparaciones', { b: true, tam: 13 }), { despues: 2 }));
  c.push(wp_(wr_('Solicitud ' + m.id + ' · enviada desde la app el ' + fechaCorta_(m.enviada.slice(0, 10)) + ' a las ' + hora, { i: true, tam: 9, color: '666666' }), { despues: 6 }));

  c.push(titulo('SECCIÓN 1. DATOS DE LA SOLICITANTE'));
  c.push(wp_([wr_('1. Fecha de solicitud: ', { b: true }), fecha(m.enviada.slice(0, 10)), wr_(hora ? ' (' + hora + ')' : ''),
              wr_('     2. Nombre de la mamá: ', { b: true }), valor(m.nombre)]));
  c.push(wp_([wr_('3. Número de apartamento: ', { b: true }), valor(m.apartamento)]));

  c.push(titulo('SECCIÓN 2. DESCRIPCIÓN DEL PROBLEMA'));
  c.push(preg('4. Describa el problema con el mayor detalle posible.'));
  c.push(wp_(wr_('¿Qué sucede? ¿Qué dejó de funcionar o qué daño presenta?'), { despues: 4, juntos: true }));
  c.push(bloque(m.descripcion));
  c.push(preg('5. ¿Cuándo ocurrió o desde cuándo presenta el problema?'));
  c.push(wp_([wr_('Fecha aproximada: '), fecha(m.fechaAproximada)], { juntos: true }));
  c.push(wp_(DESDE_MANT.map(function (t) { return casilla_(m.desdeCuando === t, t + '.') + wr_('     '); }), { sangria: 0.6 }));
  c.push(preg('6. ¿El problema está causando algún riesgo o afectación importante?'));
  c.push(wp_([casilla_(m.riesgo === 'Sí', 'Sí.'), wr_('     '), casilla_(m.riesgo === 'No', 'No.'), wr_('     '),
              wr_('Si respondió que sí, explique brevemente:', { b: true, i: true })], { sangria: 0.6, juntos: true }));
  c.push(bloque(m.riesgoDetalle));
  c.push(preg('7. ¿El problema requiere atención inmediata?'));
  c.push(wp_([casilla_(m.inmediata === 'Sí', 'Sí, existe un riesgo para las personas o el apartamento.'), wr_('     '),
              casilla_(m.inmediata === 'No', 'No, puede programarse una revisión.')]));
  c.push(wp_(wr_('Si existe un riesgo inmediato, como olor a gas, chispas eléctricas o una fuga importante de agua, debe reportarse directamente al personal responsable sin esperar a entregar esta boleta.', { i: true }), { antes: 4 }));

  c.push(titulo('SECCIÓN 3. INFORMACIÓN ADICIONAL'));
  c.push(wp_([wr_('8. ¿Hay alguna información adicional que debamos conocer?', { b: true }),
              wr_(' Por ejemplo: si el problema ocurre constantemente, si ya se intentó reportar o si sucede únicamente en determinadas circunstancias.')],
             { despues: 4, juntos: true }));
  c.push(bloque(m.infoAdicional));
  c.push(wp_(wr_('9. Firma de la solicitante:', { b: true }), { antes: 12, despues: 2, juntos: true }));
  const firmaM = imagen(firmas.solicitante, 'firma_solicitante', 4.5);
  c.push(wp_(firmaM || wr_('_______________________________'), { despues: 0, juntos: true }));
  if (firmaM) c.push(wp_(wr_('Firmada en la app por ' + m.nombre + (cedulaValida_(m.cedula || '') ? ' (cédula ' + m.cedula.replace(/^(\d)(\d{4})(\d{4})$/, '$1-$2-$3') + ')' : '') +
                             ' el ' + fechaCorta_(m.enviada.slice(0, 10)) + ' a las ' + hora + '.', { i: true, tam: 9, color: '666666' })));

  c.push(titulo('SECCIÓN 4. USO EXCLUSIVO DEL EQUIPO DE MANTENIMIENTO', { run: { color: ROJO }, par: { salto: true, antes: 0 } }));
  c.push(wp_([wr_('Esta sección debe ser ', { i: true }), wr_('completada por el personal encargado.', { i: true, b: true, u: true })], { despues: 10 }));
  c.push(wp_([wr_('Fecha de recepción: ', { b: true }), fecha(m.fechaRecepcion)]));
  c.push(preg('Prioridad asignada:'));
  [['Alta', 'Alta: requiere atención urgente.'], ['Media', 'Media: requiere atención próxima.'], ['Baja', 'Baja: puede programarse.']].forEach(function (p) {
    c.push(wp_([wr_('•  '), casilla_(m.prioridad === p[0], p[1])], { sangria: 0.6, despues: 2 }));
  });
  c.push(preg('Evaluación del problema:')); c.push(bloque(m.evaluacion));
  c.push(preg('Trabajo o reparación por realizar:')); c.push(bloque(m.trabajo));
  c.push(preg('Materiales o repuestos necesarios:')); c.push(bloque(m.materiales));
  c.push(wp_([wr_('Responsable asignado: ', { b: true }), valor(m.responsable)], { antes: 6 }));
  c.push(wp_([wr_('Fecha de atención: ', { b: true }), fecha(m.fechaAtencion)], { antes: 6 }));
  c.push(preg('Estado de la solicitud:'));
  ESTADOS_MANT.forEach(function (e) { c.push(wp_([wr_('•  '), casilla_(m.estado === e, e + '.')], { sangria: 0.6, despues: 2 })); });
  c.push(preg('Observaciones y trabajo realizado:')); c.push(bloque(m.observaciones));
  c.push(wp_([wr_('Fecha de finalización: ', { b: true }), fecha(m.fechaFinalizacion)], { antes: 6 }));
  c.push(wp_(wr_('Firma del responsable:', { b: true }), { antes: 12, despues: 2, juntos: true }));
  c.push(wp_(imagen(firmas.responsable, 'firma_responsable', 4.5) || wr_('_______________________________'), { despues: 0 }));
  if (m.actualizadaPor) c.push(wp_(wr_('Sección 4 actualizada en la app por ' + m.actualizadaPor + ' el ' + fechaCorta_(m.actualizada.slice(0, 10)) +
                                       ' a las ' + m.actualizada.slice(11, 16) + '.', { i: true, tam: 9, color: '666666' }), { antes: 6 }));

  const logo = Utilities.base64Decode(LOGO_DOCX);
  const doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>' + c.join('') +
    '<w:sectPr><w:footerReference w:type="default" r:id="rIdPie"/><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1021" w:right="1247" w:bottom="1021" w:left="1247" w:header="510" w:footer="340" w:gutter="0"/></w:sectPr></w:body></w:document>';
  const pie = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + NS + '>' +
    wp_(wimg_('rIdLogo', 1, tamPng_(logo), 2, 'logo'), { alin: 'right', despues: 0 }) + '</w:ftr>';
  const rel = function (id, tipo, destino) {
    return '<Relationship Id="' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + tipo + '" Target="' + destino + '"/>';
  };
  const relsDoc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    rel('rIdEstilos', 'styles', 'styles.xml') + rel('rIdPie', 'footer', 'footer1.xml') +
    media.map(function (x) { return rel(x.rId, 'image', 'media/' + x.nombre); }).join('') + '</Relationships>';
  const relsPie = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    rel('rIdLogo', 'image', 'media/logo.png') + '</Relationships>';
  const estilos = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>' +
    '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="es-CR"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="250" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>';
  const tipos = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';
  const relsRaiz = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>';
  const ahoraIso = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:title>' + xml_('Boleta de mantenimiento ' + m.id) + '</dc:title><dc:creator>Proyecto Bethany</dc:creator>' +
    '<dcterms:created xsi:type="dcterms:W3CDTF">' + ahoraIso + '</dcterms:created></cp:coreProperties>';
  const blob = function (contenido, nombre, tipo) { return Utilities.newBlob(contenido, tipo || 'application/xml', nombre); };
  const partes = [
    blob(tipos, '[Content_Types].xml'), blob(relsRaiz, '_rels/.rels'), blob(core, 'docProps/core.xml'),
    blob(doc, 'word/document.xml'), blob(estilos, 'word/styles.xml'), blob(pie, 'word/footer1.xml'),
    blob(relsDoc, 'word/_rels/document.xml.rels'), blob(relsPie, 'word/_rels/footer1.xml.rels'),
    blob(logo, 'word/media/logo.png', 'image/png')
  ].concat(media.map(function (x) { return blob(x.bytes, 'word/media/' + x.nombre, 'image/png'); }));
  return Utilities.zip(partes, 'boleta.docx');
}