const express = require('express');
const mysql2 = require('mysql2');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const ExcelJS = require('exceljs');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const nodemailer = require('nodemailer');
const cron = require('node-cron');
const QRCode = require('qrcode');
require('dotenv').config();

const JWT_SECRET = process.env.JWT_SECRET;
const JWT_EXPIRA = process.env.JWT_EXPIRA || '8h';

if (!JWT_SECRET) {
    console.error('Falta JWT_SECRET en el archivo .env (ver instrucciones de instalación).');
    process.exit(1);
}

const app = express();
app.use(cors());
app.use(express.json());
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

const carpetaImagenesEquipos = path.join(__dirname, 'uploads', 'equipos');
fs.mkdirSync(carpetaImagenesEquipos, { recursive: true });

const uploadImportacion = multer({ storage: multer.memoryStorage(), limits: { fileSize: 15 * 1024 * 1024 } }).single('archivo');
const uploadImagenEquipo = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } }).single('imagen');

// Configuración de la conexión a MySQL en XAMPP
const db = mysql2.createConnection({
    host: 'localhost',
    user: 'root',
    password: '',
    database: 'biomant'
});

db.connect((err) => {
    if (err) {
        console.error('Error al conectar a la base de datos:', err);
        return;
    }
    console.log('¡Conectado exitosamente a la base de datos de XAMPP!');
});

// Panel visual de inventario (el propio index.html redirige al login si no hay sesión)
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Vista de login
app.get('/login', (req, res) => {
    res.sendFile(path.join(__dirname, 'login.html'));
});

// Iniciar servidor
const PORT = 3000;

// =============================================================================
// AUTENTICACIÓN Y ROLES
// =============================================================================

const ROLES = ['Administrador', 'Tecnico'];

// Hash de relleno: permite comparar siempre contra un hash, exista o no el
// correo, para que el tiempo de respuesta no revele qué correos están registrados.
const HASH_RELLENO = bcrypt.hashSync('relleno-no-es-una-clave-real', 10);

// Exige una sesión válida (cabecera Authorization: Bearer <token>)
function verificarToken(req, res, next) {
    const cabecera = req.headers.authorization || '';
    const [esquema, token] = cabecera.split(' ');

    if (esquema !== 'Bearer' || !token) {
        res.status(401).json({ error: 'Debes iniciar sesión' });
        return;
    }

    jwt.verify(token, JWT_SECRET, (err, payload) => {
        if (err) {
            res.status(401).json({ error: 'Sesión inválida o vencida' });
            return;
        }

        req.usuario = {
            id_usuario: payload.id_usuario,
            nombre: payload.nombre,
            correo: payload.correo,
            rol: payload.rol
        };
        next();
    });
}

// Exige además rol Administrador. Siempre se usa DESPUÉS de verificarToken.
function soloAdministrador(req, res, next) {
    if (!req.usuario || req.usuario.rol !== 'Administrador') {
        res.status(403).json({ error: 'Esta acción es solo para el rol Administrador' });
        return;
    }
    next();
}

const CAMPOS_IMPORTACION = {
    id_equipo: ['id', 'idequipo', 'codigo', 'codigoequipo'],
    nombre_equipo: ['nombre', 'nombreequipo', 'equipo'],
    marca: ['marca', 'fabricante'],
    modelo: ['modelo'],
    numero_serie: ['serie', 'numeroserie', 'numeroserial', 'serial'],
    nivel_riesgo: ['riesgo', 'nivelriesgo', 'criticidad'],
    ubicacion: ['ubicacion', 'departamento', 'area', 'servicio'],
    estado: ['estado', 'estadoactual'],
    tipo_categoria: ['categoria', 'tipocategoria', 'tipo'],
    proximo_mto_preventivo: ['proximomantenimiento', 'fechaproximomantenimiento', 'proximomantenimientopreventivo', 'proximomtopreventivo', 'fechaproximomtopreventivo'],
    foto_equipo: ['foto', 'fotografia', 'fotoequipo', 'imagenequipo']
};
const CAMPOS_IMPORTACION_REQUERIDOS = ['id_equipo', 'nombre_equipo', 'marca', 'modelo', 'numero_serie', 'nivel_riesgo', 'ubicacion', 'estado', 'tipo_categoria'];

function normalizarEncabezado(valor) {
    return String(valor ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function mapearEncabezados(encabezados) {
    const campos = encabezados.map((encabezado) => {
        const clave = normalizarEncabezado(encabezado);
        return Object.entries(CAMPOS_IMPORTACION).find(([, alias]) => alias.includes(clave))?.[0] || null;
    });
    const faltantes = CAMPOS_IMPORTACION_REQUERIDOS.filter((campo) => !campos.includes(campo));
    return { campos, faltantes };
}

function valorCeldaExcel(valor) {
    if (valor instanceof Date) return valor.toISOString().slice(0, 10);
    if (valor && typeof valor === 'object') {
        if (valor.text) return String(valor.text).trim();
        if (valor.result !== undefined) return String(valor.result).trim();
        return '';
    }
    return String(valor ?? '').trim();
}

async function extraerFilasExcel(buffer) {
    const libro = new ExcelJS.Workbook();
    await libro.xlsx.load(buffer);
    const hoja = libro.worksheets[0];
    if (!hoja || hoja.rowCount < 2) throw new Error('El Excel debe incluir encabezados y al menos una fila de datos.');
    const encabezados = hoja.getRow(1).values.slice(1).map(valorCeldaExcel);
    const { campos, faltantes } = mapearEncabezados(encabezados);
    if (faltantes.length) throw new Error(`Faltan columnas requeridas: ${faltantes.join(', ')}.`);
    const filas = [];
    hoja.eachRow({ includeEmpty: false }, (fila, numero) => {
        if (numero === 1) return;
        const valores = fila.values.slice(1);
        if (!valores.some((valor) => valor !== null && valor !== undefined && valor !== '')) return;
        const registro = {};
        campos.forEach((campo, indice) => { if (campo) registro[campo] = valorCeldaExcel(valores[indice]); });
        filas.push({ numero, registro });
    });
    return filas;
}

async function extraerTextoPDF(buffer) {
    if (typeof pdfParse === 'function') return (await pdfParse(buffer)).text || '';
    if (pdfParse.PDFParse) {
        const parser = new pdfParse.PDFParse({ data: buffer });
        try {
            const resultado = await parser.getText();
            return resultado.text || '';
        } finally {
            await parser.destroy();
        }
    }
    throw new Error('La versión instalada de pdf-parse no tiene una API compatible.');
}

async function extraerFilasPDF(buffer) {
    const texto = await extraerTextoPDF(buffer);
    const lineas = texto.split(/\r?\n/).map((linea) => linea.trim()).filter(Boolean);
    const separar = (linea) => linea.split(/\s*\|\s*|\t+|;|\s{2,}/).map((valor) => valor.trim()).filter(Boolean);
    const indiceEncabezado = lineas.findIndex((linea) => {
        const mapping = mapearEncabezados(separar(linea));
        return mapping.campos.includes('id_equipo') && mapping.campos.includes('nombre_equipo');
    });
    if (indiceEncabezado < 0) throw new Error('No se encontró una tabla de texto con encabezados. Usa columnas separadas por |, punto y coma, tabulador o espacios múltiples; los PDF escaneados requieren OCR.');
    const encabezados = separar(lineas[indiceEncabezado]);
    const { campos, faltantes } = mapearEncabezados(encabezados);
    if (faltantes.length) throw new Error(`Faltan columnas requeridas: ${faltantes.join(', ')}.`);
    return lineas.slice(indiceEncabezado + 1).map((linea, indice) => {
        const valores = separar(linea);
        const registro = {};
        campos.forEach((campo, columna) => { if (campo) registro[campo] = valores[columna] || ''; });
        return { numero: indiceEncabezado + indice + 2, registro };
    }).filter(({ registro }) => Object.values(registro).some(Boolean));
}

// Iniciar sesión: devuelve un token JWT y los datos básicos del usuario
app.post('/api/login', (req, res) => {
    const { correo, password } = req.body || {};

    if (!correo || !password) {
        res.status(400).json({ error: 'Ingresa tu correo y tu contraseña' });
        return;
    }

    const query = 'SELECT id_usuario, nombre, correo, password, rol FROM usuarios WHERE correo = ?';

    db.query(query, [String(correo).trim().toLowerCase()], (err, results) => {
        if (err) {
            console.error('Error al consultar el usuario:', err);
            res.status(500).json({ error: 'Error al iniciar sesión' });
            return;
        }

        const encontrado = results[0];
        const hash = encontrado ? encontrado.password : HASH_RELLENO;

        bcrypt.compare(String(password), hash, (errHash, coincide) => {
            if (errHash) {
                console.error('Error al verificar la contraseña:', errHash);
                res.status(500).json({ error: 'Error al iniciar sesión' });
                return;
            }

            if (!encontrado || !coincide) {
                res.status(401).json({ error: 'Correo o contraseña incorrectos' });
                return;
            }

            const usuario = {
                id_usuario: encontrado.id_usuario,
                nombre: encontrado.nombre,
                correo: encontrado.correo,
                rol: encontrado.rol
            };

            const token = jwt.sign(usuario, JWT_SECRET, { expiresIn: JWT_EXPIRA });
            res.json({ token, usuario });
        });
    });
});

// Listar usuarios (solo Administrador). Nunca devuelve el hash.
app.get('/api/usuarios', verificarToken, soloAdministrador, (req, res) => {
    const query = 'SELECT id_usuario, nombre, correo, rol, creado_en FROM usuarios ORDER BY nombre';

    db.query(query, (err, results) => {
        if (err) {
            console.error('Error al listar usuarios:', err);
            res.status(500).json({ error: 'Error al obtener los usuarios' });
            return;
        }
        res.json(results);
    });
});

// Crear un usuario (solo Administrador). La contraseña se guarda con bcrypt.
app.post('/api/usuarios', verificarToken, soloAdministrador, (req, res) => {
    const { nombre, correo, password, rol } = req.body || {};
    const faltantes = ['nombre', 'correo', 'password', 'rol'].filter(
        (campo) => !req.body || !String(req.body[campo] ?? '').trim()
    );

    if (faltantes.length > 0) {
        res.status(400).json({ error: 'Faltan datos requeridos', campos_faltantes: faltantes });
        return;
    }

    if (!ROLES.includes(rol)) {
        res.status(400).json({ error: `El rol debe ser uno de: ${ROLES.join(', ')}` });
        return;
    }

    if (String(password).length < 8) {
        res.status(400).json({ error: 'La contraseña debe tener al menos 8 caracteres' });
        return;
    }

    bcrypt.hash(String(password), 10, (errHash, hash) => {
        if (errHash) {
            console.error('Error al cifrar la contraseña:', errHash);
            res.status(500).json({ error: 'Error al registrar el usuario' });
            return;
        }

        const query = 'INSERT INTO usuarios (nombre, correo, password, rol) VALUES (?, ?, ?, ?)';
        const valores = [String(nombre).trim(), String(correo).trim().toLowerCase(), hash, rol];

        db.query(query, valores, (err, result) => {
            if (err) {
                if (err.code === 'ER_DUP_ENTRY') {
                    res.status(409).json({ error: 'Ya existe un usuario con ese correo' });
                    return;
                }
                console.error('Error al crear el usuario:', err);
                res.status(500).json({ error: 'Error al registrar el usuario' });
                return;
            }

            res.status(201).json({
                mensaje: 'Usuario registrado correctamente',
                id_usuario: result.insertId
            });
        });
    });
});

// Eliminar usuario (solo Administrador). Protege la cuenta propia y al último administrador.
app.delete('/api/usuarios/:id', verificarToken, soloAdministrador, (req, res) => {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id < 1) {
        res.status(400).json({ error: 'El identificador del usuario no es válido' });
        return;
    }
    if (id === Number(req.usuario.id_usuario)) {
        res.status(400).json({ error: 'No puedes eliminar tu propia cuenta mientras tienes la sesión iniciada' });
        return;
    }

    db.query('SELECT rol FROM usuarios WHERE id_usuario = ?', [id], (errUsuario, usuariosEncontrados) => {
        if (errUsuario) {
            console.error('Error al consultar el usuario antes de eliminarlo:', errUsuario);
            res.status(500).json({ error: 'Error al eliminar el usuario' });
            return;
        }
        if (!usuariosEncontrados.length) {
            res.status(404).json({ error: 'Usuario no encontrado' });
            return;
        }
        const eliminarUsuario = () => db.query('DELETE FROM usuarios WHERE id_usuario = ?', [id], (errEliminar, resultado) => {
            if (errEliminar) {
                console.error('Error al eliminar el usuario:', errEliminar);
                res.status(500).json({ error: 'No se pudo eliminar el usuario. Puede tener registros asociados.' });
                return;
            }
            if (!resultado.affectedRows) {
                res.status(404).json({ error: 'Usuario no encontrado' });
                return;
            }
            res.json({ mensaje: 'Usuario eliminado correctamente', id_usuario: id });
        });

        if (usuariosEncontrados[0].rol !== 'Administrador') {
            eliminarUsuario();
            return;
        }
        db.query("SELECT COUNT(*) AS cantidad FROM usuarios WHERE rol = 'Administrador'", (errConteo, filas) => {
            if (errConteo) {
                console.error('Error al verificar administradores:', errConteo);
                res.status(500).json({ error: 'Error al verificar administradores' });
                return;
            }
            if (Number(filas[0].cantidad) <= 1) {
                res.status(400).json({ error: 'No se puede eliminar el último administrador del sistema' });
                return;
            }
            eliminarUsuario();
        });
    });
});

// =============================================================================
// INVENTARIO DE EQUIPOS
// =============================================================================

// Todo /api/equipos requiere sesión iniciada (cualquier rol).
// Las acciones que modifican el inventario además exigen rol Administrador.
app.use('/api/equipos', verificarToken);

const SELECT_EQUIPO = `
    SELECT e.*,
           DATE_FORMAT(e.proximo_mto_preventivo, '%Y-%m-%d') AS fecha_proximo_mto_preventivo,
           DATEDIFF(e.proximo_mto_preventivo, CURRENT_DATE) AS dias_hasta_preventivo,
           CASE
             WHEN e.proximo_mto_preventivo IS NULL THEN 'Sin programar'
             WHEN e.proximo_mto_preventivo < CURRENT_DATE THEN 'Vencido'
             WHEN e.proximo_mto_preventivo <= DATE_ADD(CURRENT_DATE, INTERVAL 15 DAY) THEN 'Próximo'
             ELSE 'Al día'
           END AS estado_mantenimiento_preventivo
    FROM equipos e`;

// Campos requeridos para crear o actualizar un equipo
const CAMPOS_EQUIPO = [
    'id_equipo',
    'nombre_equipo',
    'marca',
    'modelo',
    'numero_serie',
    'nivel_riesgo',
    'ubicacion',
    'estado',
    'tipo_categoria',
    'proximo_mto_preventivo'
];

const CAMPOS_ACTUALIZACION = CAMPOS_EQUIPO.filter((campo) => campo !== 'id_equipo');

function validarCamposRequeridos(datos, campos) {
    return campos.filter((campo) => {
        const valor = datos[campo];
        return valor === undefined || valor === null || String(valor).trim() === '';
    });
}

function fechaISOValidaOpcional(fecha) {
    if (fecha === undefined || fecha === null || fecha === '') return true;
    return typeof fecha === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(fecha) && !Number.isNaN(Date.parse(`${fecha}T00:00:00Z`)) && new Date(`${fecha}T00:00:00Z`).toISOString().slice(0, 10) === fecha;
}

// Ruta para obtener todos los equipos del inventario
app.get('/api/equipos', (req, res) => {
    const query = SELECT_EQUIPO;
    db.query(query, (err, results) => {
        if (err) {
            console.error('Error al consultar los equipos:', err);
            res.status(500).json({ error: 'Error al obtener el inventario' });
            return;
        }
        res.json(results);
    });
});

// Filtrar equipos por categoría (Biomédico o Tecnología)
// Se declara antes de /:id para evitar que "buscar" se interprete como identificador
app.get('/api/equipos/buscar/categoria/:tipo', (req, res) => {
    const { tipo } = req.params;
    const query = `${SELECT_EQUIPO} WHERE e.tipo_categoria = ?`;

    db.query(query, [tipo], (err, results) => {
        if (err) {
            console.error('Error al filtrar equipos por categoría:', err);
            res.status(500).json({ error: 'Error al buscar equipos por categoría' });
            return;
        }
        res.json(results);
    });
});

// Obtener un equipo específico por id_equipo
app.get('/api/equipos/:id', (req, res) => {
    const { id } = req.params;
    const query = `${SELECT_EQUIPO} WHERE e.id_equipo = ?`;

    db.query(query, [id], (err, results) => {
        if (err) {
            console.error('Error al consultar el equipo:', err);
            res.status(500).json({ error: 'Error al obtener el equipo' });
            return;
        }

        if (results.length === 0) {
            res.status(404).json({ error: 'Equipo no encontrado' });
            return;
        }

        res.json(results[0]);
    });
});

// Crear un nuevo equipo en el inventario (solo Administrador)
app.post('/api/equipos', soloAdministrador, (req, res) => {
    const faltantes = validarCamposRequeridos(req.body, CAMPOS_EQUIPO.filter((campo) => campo !== 'proximo_mto_preventivo'));

    if (faltantes.length > 0) {
        res.status(400).json({
            error: 'Faltan datos requeridos',
            campos_faltantes: faltantes
        });
        return;
    }

    if (!fechaISOValidaOpcional(req.body.proximo_mto_preventivo)) {
        res.status(400).json({ error: 'La fecha del próximo mantenimiento preventivo no es válida' });
        return;
    }

    const valores = CAMPOS_EQUIPO.map((campo) => campo === 'proximo_mto_preventivo' ? (req.body[campo] || null) : req.body[campo]);
    const placeholders = CAMPOS_EQUIPO.map(() => '?').join(', ');
    const query = `INSERT INTO equipos (${CAMPOS_EQUIPO.join(', ')}) VALUES (${placeholders})`;

    db.query(query, valores, (err) => {
        if (err) {
            console.error('Error al crear el equipo:', err);
            res.status(500).json({ error: 'Error al registrar el equipo' });
            return;
        }

        res.status(201).json({
            mensaje: 'Equipo registrado correctamente',
            id_equipo: req.body.id_equipo
        });
    });
});

// Actualizar los datos de un equipo existente (solo Administrador)
app.put('/api/equipos/:id', soloAdministrador, (req, res) => {
    const { id } = req.params;
    const camposObligatorios = CAMPOS_ACTUALIZACION.filter((campo) => campo !== 'proximo_mto_preventivo');
    const faltantes = validarCamposRequeridos(req.body, camposObligatorios);

    if (faltantes.length > 0) {
        res.status(400).json({
            error: 'Faltan datos requeridos',
            campos_faltantes: faltantes
        });
        return;
    }

    if (!fechaISOValidaOpcional(req.body.proximo_mto_preventivo)) {
        res.status(400).json({ error: 'La fecha del próximo mantenimiento preventivo no es válida' });
        return;
    }

    const setClause = CAMPOS_ACTUALIZACION.map((campo) => `${campo} = ?`).join(', ');
    const valores = [...CAMPOS_ACTUALIZACION.map((campo) => campo === 'proximo_mto_preventivo' ? (req.body[campo] || null) : req.body[campo]), id];
    const query = `UPDATE equipos SET ${setClause} WHERE id_equipo = ?`;

    db.query(query, valores, (err, result) => {
        if (err) {
            console.error('Error al actualizar el equipo:', err);
            res.status(500).json({ error: 'Error al actualizar el equipo' });
            return;
        }

        if (result.affectedRows === 0) {
            res.status(404).json({ error: 'Equipo no encontrado' });
            return;
        }

        res.status(200).json({
            mensaje: 'Equipo actualizado correctamente',
            id_equipo: id
        });
    });
});

// Eliminar un equipo por su id_equipo (solo Administrador)
app.delete('/api/equipos/:id', soloAdministrador, (req, res) => {
    const { id } = req.params;
    const query = 'DELETE FROM equipos WHERE id_equipo = ?';

    db.query(query, [id], (err, result) => {
        if (err) {
            console.error('Error al eliminar el equipo:', err);
            res.status(500).json({ error: 'Error al eliminar el equipo' });
            return;
        }

        if (result.affectedRows === 0) {
            res.status(404).json({ error: 'Equipo no encontrado' });
            return;
        }

        res.status(200).json({
            mensaje: 'Equipo eliminado correctamente',
            id_equipo: id
        });
    });
});

// -----------------------------------------------------------------------------
// ESTADO DEL EQUIPO
// Administrador: cualquier estado. Técnico: solo alterna entre Operativo y
// En mantenimiento; dar de baja o reactivar un equipo fuera de servicio es del Administrador.
// -----------------------------------------------------------------------------
const ESTADOS_EQUIPO = ['Operativo', 'En mantenimiento', 'Fuera de servicio'];

app.patch('/api/equipos/:id/estado', (req, res) => {
    const { id } = req.params;
    const pedido = String((req.body && req.body.estado) ?? '').trim().toLowerCase();
    const nuevo = ESTADOS_EQUIPO.find((estado) => estado.toLowerCase() === pedido);

    if (!nuevo) {
        res.status(400).json({ error: `El estado debe ser uno de: ${ESTADOS_EQUIPO.join(', ')}` });
        return;
    }

    db.query('SELECT estado FROM equipos WHERE id_equipo = ?', [id], (err, results) => {
        if (err) {
            console.error('Error al consultar el estado del equipo:', err);
            res.status(500).json({ error: 'Error al cambiar el estado' });
            return;
        }

        if (results.length === 0) {
            res.status(404).json({ error: 'Equipo no encontrado' });
            return;
        }

        const actual = String(results[0].estado || '').toLowerCase();
        const afectaBaja = actual === 'fuera de servicio' || nuevo === 'Fuera de servicio';

        if (req.usuario.rol !== 'Administrador' && afectaBaja) {
            res.status(403).json({ error: 'Solo un Administrador puede dar de baja o reactivar un equipo' });
            return;
        }

        db.query('UPDATE equipos SET estado = ? WHERE id_equipo = ?', [nuevo, id], (errUpdate) => {
            if (errUpdate) {
                console.error('Error al cambiar el estado del equipo:', errUpdate);
                res.status(500).json({ error: 'Error al cambiar el estado' });
                return;
            }
            res.json({ mensaje: 'Estado actualizado correctamente', id_equipo: id, estado: nuevo });
        });
    });
});

// Programar o limpiar la fecha preventiva próxima (solo Administrador).
app.patch('/api/equipos/:id/proximo-preventivo', soloAdministrador, (req, res) => {
    const fecha = req.body && req.body.fecha_proximo_mto_preventivo;
    if (fecha !== null && fecha !== '' && (typeof fecha !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(fecha) || Number.isNaN(Date.parse(`${fecha}T00:00:00Z`)) || new Date(`${fecha}T00:00:00Z`).toISOString().slice(0, 10) !== fecha)) {
        res.status(400).json({ error: 'La fecha programada no es válida' });
        return;
    }
    db.query('UPDATE equipos SET proximo_mto_preventivo = ? WHERE id_equipo = ?', [fecha || null, req.params.id], (err, result) => {
        if (err) {
            console.error('Error al actualizar la fecha preventiva:', err);
            res.status(500).json({ error: 'Error al actualizar la fecha preventiva. Verifica que ejecutaste alertas_analitica.sql.' });
            return;
        }
        if (!result.affectedRows) {
            res.status(404).json({ error: 'Equipo no encontrado' });
            return;
        }
        db.query(`${SELECT_EQUIPO} WHERE e.id_equipo = ?`, [req.params.id], (errEquipo, equipos) => {
            if (errEquipo) {
                res.status(500).json({ error: 'La fecha se guardó, pero no se pudo leer el equipo' });
                return;
            }
            res.json(equipos[0]);
        });
    });
});

app.get('/api/equipos/:id/qr', async (req, res) => {
    db.query('SELECT id_equipo FROM equipos WHERE id_equipo = ?', [req.params.id], async (err, equipos) => {
        if (err) {
            res.status(500).json({ error: 'No se pudo generar el QR.' });
            return;
        }
        if (!equipos.length) {
            res.status(404).json({ error: 'Equipo no encontrado.' });
            return;
        }
        try {
            const baseUrl = (process.env.APP_PUBLIC_URL || `${req.protocol}://${req.get('host')}`).replace(/\/$/, '');
            const fichaUrl = `${baseUrl}/?equipo=${encodeURIComponent(equipos[0].id_equipo)}`;
            const qr = await QRCode.toDataURL(fichaUrl, { width: 320, margin: 2, errorCorrectionLevel: 'M' });
            res.json({ url: fichaUrl, qr });
        } catch (error) {
            console.error('Error al generar el QR:', error);
            res.status(500).json({ error: 'No se pudo generar el QR.' });
        }
    });
});

// Carga de foto: solo Administrador, formatos JPEG/PNG/WebP, máximo 5 MB.
app.post('/api/equipos/:id/imagen', soloAdministrador, (req, res) => {
    uploadImagenEquipo(req, res, (errUpload) => {
        if (errUpload) {
            res.status(400).json({ error: errUpload.code === 'LIMIT_FILE_SIZE' ? 'La imagen supera el límite de 5 MB.' : 'No se pudo procesar la imagen.' });
            return;
        }
        if (!req.file) {
            res.status(400).json({ error: 'Selecciona una imagen JPEG, PNG o WebP.' });
            return;
        }
        const firma = req.file.buffer;
        let extension = null;
        if (firma.length >= 3 && firma[0] === 0xff && firma[1] === 0xd8 && firma[2] === 0xff) extension = 'jpg';
        else if (firma.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))) extension = 'png';
        else if (firma.length >= 12 && firma.toString('ascii', 0, 4) === 'RIFF' && firma.toString('ascii', 8, 12) === 'WEBP') extension = 'webp';
        if (!extension) {
            res.status(400).json({ error: 'El archivo no parece ser una imagen JPEG, PNG o WebP válida.' });
            return;
        }
        const nombreArchivo = `${crypto.randomUUID()}.${extension}`;
        const rutaArchivo = path.join(carpetaImagenesEquipos, nombreArchivo);
        const rutaPublica = `/uploads/equipos/${nombreArchivo}`;
        fs.writeFile(rutaArchivo, firma, (errWrite) => {
            if (errWrite) {
                console.error('Error al guardar la imagen:', errWrite);
                res.status(500).json({ error: 'No se pudo guardar la imagen.' });
                return;
            }
            db.query('UPDATE equipos SET foto_equipo = ? WHERE id_equipo = ?', [rutaPublica, req.params.id], (errUpdate, resultado) => {
                if (errUpdate || !resultado.affectedRows) {
                    fs.unlink(rutaArchivo, () => {});
                    if (errUpdate) console.error('Error al asociar la imagen al equipo:', errUpdate);
                    res.status(errUpdate ? 500 : 404).json({ error: errUpdate ? 'No se pudo asociar la imagen al equipo.' : 'Equipo no encontrado.' });
                    return;
                }
                res.json({ foto_equipo: rutaPublica });
            });
        });
    });
});

// Importar inventario desde Excel o PDF tabular (solo Administrador).
app.post('/api/equipos/importar', soloAdministrador, (req, res) => {
    uploadImportacion(req, res, async (errUpload) => {
        if (errUpload) {
            res.status(400).json({ error: errUpload.code === 'LIMIT_FILE_SIZE' ? 'El archivo supera el límite de 15 MB.' : 'No se pudo procesar el archivo.' });
            return;
        }
        if (!req.file) {
            res.status(400).json({ error: 'Selecciona un archivo .xlsx o .pdf.' });
            return;
        }

        try {
            const extension = path.extname(req.file.originalname).toLowerCase();
            if (!['.xlsx', '.pdf'].includes(extension)) {
                res.status(400).json({ error: 'El formato debe ser Excel .xlsx o PDF.' });
                return;
            }
            const filas = extension === '.xlsx'
                ? await extraerFilasExcel(req.file.buffer)
                : await extraerFilasPDF(req.file.buffer);
            if (!filas.length) {
                res.status(400).json({ error: 'El archivo no contiene filas de equipos.' });
                return;
            }
            if (filas.length > 5000) {
                res.status(400).json({ error: 'El archivo supera el máximo de 5.000 filas por importación.' });
                return;
            }

            const insertSql = `INSERT IGNORE INTO equipos
                (id_equipo, nombre_equipo, marca, modelo, numero_serie, nivel_riesgo, ubicacion, estado, tipo_categoria, proximo_mto_preventivo, foto_equipo)
                VALUES ?`;
            const estadosValidos = ['Operativo', 'En mantenimiento', 'Fuera de servicio'];
            const categoriasValidas = ['Biomédico', 'Tecnología'];
            const riesgosValidos = ['I', 'IIA', 'IIB', 'III'];
            const errores = [];
            const registrosValidos = [];
            for (const fila of filas) {
                const registro = Object.fromEntries(Object.entries(fila.registro).map(([campo, valor]) => [campo, String(valor ?? '').trim()]));
                const faltantes = CAMPOS_IMPORTACION_REQUERIDOS.filter((campo) => !registro[campo]);
                const errorFila = (mensaje) => errores.push({ fila: fila.numero, error: mensaje });
                if (faltantes.length) { errorFila(`Faltan datos: ${faltantes.join(', ')}`); continue; }
                if (registro.id_equipo.length > 30 || registro.nombre_equipo.length > 255 || registro.marca.length > 100 || registro.modelo.length > 100 || registro.numero_serie.length > 100 || registro.ubicacion.length > 150 || (registro.foto_equipo && registro.foto_equipo.length > 255)) {
                    errorFila('Uno o más campos exceden la longitud permitida.'); continue;
                }
                if (registro.foto_equipo && !/^(https?:\/\/|\/uploads\/)/i.test(registro.foto_equipo)) {
                    errorFila('La fotografía debe usar una URL http(s) o una ruta /uploads/.'); continue;
                }
                const estado = estadosValidos.find((valor) => normalizarEncabezado(valor) === normalizarEncabezado(registro.estado));
                const categoria = categoriasValidas.find((valor) => normalizarEncabezado(valor) === normalizarEncabezado(registro.tipo_categoria));
                const riesgo = riesgosValidos.find((valor) => valor.toLowerCase() === registro.nivel_riesgo.toLowerCase());
                if (!estado || !categoria || !riesgo) { errorFila('Estado, categoría o nivel de riesgo no válido.'); continue; }
                const fechaProxima = registro.proximo_mto_preventivo || null;
                if (!fechaISOValidaOpcional(fechaProxima)) { errorFila('La fecha próxima debe tener formato AAAA-MM-DD.'); continue; }
                registrosValidos.push({ numero: fila.numero, valores: [
                    registro.id_equipo, registro.nombre_equipo, registro.marca, registro.modelo,
                    registro.numero_serie, riesgo, registro.ubicacion, estado, categoria,
                    fechaProxima, registro.foto_equipo || null
                ] });
            }
            let insertados = 0;
            let duplicados = 0;
            for (let inicio = 0; inicio < registrosValidos.length; inicio += 250) {
                const lote = registrosValidos.slice(inicio, inicio + 250);
                try {
                    const [resultado] = await db.promise().query(insertSql, [lote.map((fila) => fila.valores)]);
                    insertados += resultado.affectedRows;
                    duplicados += lote.length - resultado.affectedRows;
                } catch (errLote) {
                    console.error('Error al guardar un lote de inventario:', errLote);
                    lote.forEach((fila) => errores.push({ fila: fila.numero, error: 'La fila no se pudo guardar en la base de datos.' }));
                }
            }
            res.json({ total: filas.length, insertados, duplicados, rechazados: errores.length, errores: errores.slice(0, 50) });
        } catch (error) {
            console.error('Error al importar el inventario:', error);
            res.status(400).json({ error: error.message || 'No se pudo leer el archivo. Verifica su formato.' });
        }
    });
});

// =============================================================================
// MANTENIMIENTOS
// Cualquier usuario con sesión consulta y registra. Administradores y autores
// pueden editar o eliminar el registro.
// =============================================================================

app.use('/api/mantenimientos', verificarToken);

app.get('/api/analitica/dashboard', verificarToken, (req, res) => {
    const estadosSql = 'SELECT estado, COUNT(*) AS cantidad FROM equipos GROUP BY estado ORDER BY estado';
    const mantenimientosSql = 'SELECT tipo, COUNT(*) AS cantidad FROM mantenimientos GROUP BY tipo ORDER BY tipo';
    db.query(estadosSql, (errEstados, estados) => {
        if (errEstados) {
            console.error('Error al consultar estados para el panel:', errEstados);
            res.status(500).json({ error: 'No se pudo cargar la distribución por estado' });
            return;
        }
        db.query(mantenimientosSql, (errMantenimientos, mantenimientos) => {
            if (errMantenimientos) {
                console.error('Error al consultar mantenimientos para el panel:', errMantenimientos);
                res.status(500).json({ error: 'No se pudo cargar la frecuencia de mantenimientos' });
                return;
            }
            res.json({ equiposPorEstado: estados, mantenimientosPorTipo: mantenimientos });
        });
    });
});

const TIPOS_MANTENIMIENTO = ['Preventivo', 'Correctivo'];

const SELECT_MANTENIMIENTO = `
    SELECT m.id_mantenimiento, m.id_equipo, e.nombre_equipo, m.tipo,
           DATE_FORMAT(m.fecha, '%Y-%m-%d') AS fecha, m.detalle,
           m.id_usuario, u.nombre AS registrado_por
    FROM mantenimientos m
    JOIN equipos e ON e.id_equipo = m.id_equipo
    LEFT JOIN usuarios u ON u.id_usuario = m.id_usuario`;

function fechaHoy() {
    const ahora = new Date();
    return new Date(ahora.getTime() - ahora.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
}

// Devuelve un mensaje de error, o null si la fecha (AAAA-MM-DD) es válida y no es futura
function errorDeFecha(fecha) {
    if (typeof fecha !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
        return 'La fecha debe tener el formato AAAA-MM-DD';
    }
    const parseada = new Date(`${fecha}T00:00:00Z`);
    if (Number.isNaN(parseada.getTime()) || parseada.toISOString().slice(0, 10) !== fecha) {
        return 'La fecha no es válida';
    }
    if (fecha < '2000-01-01') return 'La fecha es demasiado antigua';
    if (fecha > fechaHoy()) return 'La fecha no puede ser futura';
    return null;
}

function obtenerMantenimiento(id, callback) {
    db.query(`${SELECT_MANTENIMIENTO} WHERE m.id_mantenimiento = ?`, [id], (err, results) => {
        callback(err, results && results[0]);
    });
}

// Listar todo el historial, del más reciente al más antiguo
app.get('/api/mantenimientos', (req, res) => {
    db.query(`${SELECT_MANTENIMIENTO} ORDER BY m.fecha DESC, m.id_mantenimiento DESC`, (err, results) => {
        if (err) {
            console.error('Error al listar mantenimientos:', err);
            res.status(500).json({ error: 'Error al obtener el historial' });
            return;
        }
        res.json(results);
    });
});

// Registrar un mantenimiento nuevo; responde con la fila completa para pintarla al instante
app.post('/api/mantenimientos', (req, res) => {
    const { id_equipo, tipo, fecha } = req.body || {};
    const detalle = String((req.body && req.body.detalle) ?? '').trim();

    if (!id_equipo || !tipo || !fecha || !detalle) {
        res.status(400).json({ error: 'Faltan datos requeridos (equipo, tipo, fecha y detalle)' });
        return;
    }
    if (!TIPOS_MANTENIMIENTO.includes(tipo)) {
        res.status(400).json({ error: `El tipo debe ser: ${TIPOS_MANTENIMIENTO.join(' o ')}` });
        return;
    }
    const errorFecha = errorDeFecha(fecha);
    if (errorFecha) {
        res.status(400).json({ error: errorFecha });
        return;
    }
    if (detalle.length < 10 || detalle.length > 2000) {
        res.status(400).json({ error: 'El detalle debe tener entre 10 y 2000 caracteres' });
        return;
    }

    db.query('SELECT 1 FROM equipos WHERE id_equipo = ?', [id_equipo], (err, existe) => {
        if (err) {
            console.error('Error al validar el equipo:', err);
            res.status(500).json({ error: 'Error al registrar el mantenimiento' });
            return;
        }
        if (existe.length === 0) {
            res.status(404).json({ error: 'El equipo no existe' });
            return;
        }

        const insertar = 'INSERT INTO mantenimientos (id_equipo, tipo, fecha, detalle, id_usuario) VALUES (?, ?, ?, ?, ?)';
        db.query(insertar, [id_equipo, tipo, fecha, detalle, req.usuario.id_usuario], (errInsert, result) => {
            if (errInsert) {
                console.error('Error al registrar el mantenimiento:', errInsert);
                res.status(500).json({ error: 'Error al registrar el mantenimiento' });
                return;
            }
            obtenerMantenimiento(result.insertId, (errFila, fila) => {
                if (errFila || !fila) {
                    console.error('Error al leer el mantenimiento creado:', errFila);
                    res.status(500).json({ error: 'El mantenimiento se guardó, pero no se pudo leer' });
                    return;
                }
                res.status(201).json(fila);
            });
        });
    });
});

function autorizarEdicionMantenimiento(id, req, res, callback) {
    db.query('SELECT id_usuario FROM mantenimientos WHERE id_mantenimiento = ?', [id], (err, results) => {
        if (err) {
            console.error('Error al consultar el mantenimiento:', err);
            res.status(500).json({ error: 'Error al consultar el mantenimiento' });
            return;
        }
        if (results.length === 0) {
            res.status(404).json({ error: 'Mantenimiento no encontrado' });
            return;
        }
        if (req.usuario.rol !== 'Administrador' && results[0].id_usuario !== req.usuario.id_usuario) {
            res.status(403).json({ error: 'Solo puedes modificar los mantenimientos que tú registraste' });
            return;
        }
        callback();
    });
}

// Editar todos los campos del mantenimiento
app.put('/api/mantenimientos/:id', (req, res) => {
    const { id } = req.params;
    const { id_equipo, tipo, fecha } = req.body || {};
    const detalle = String((req.body && req.body.detalle) ?? '').trim();
    const errorFecha = errorDeFecha(fecha);
    if (!/^\d+$/.test(id) || !id_equipo || !tipo || !fecha || !detalle) {
        res.status(400).json({ error: 'Faltan datos requeridos (equipo, tipo, fecha y detalle)' });
        return;
    }
    if (!TIPOS_MANTENIMIENTO.includes(tipo)) {
        res.status(400).json({ error: `El tipo debe ser: ${TIPOS_MANTENIMIENTO.join(' o ')}` });
        return;
    }
    if (errorFecha) {
        res.status(400).json({ error: errorFecha });
        return;
    }
    if (detalle.length < 10 || detalle.length > 2000) {
        res.status(400).json({ error: 'El detalle debe tener entre 10 y 2000 caracteres' });
        return;
    }

    autorizarEdicionMantenimiento(id, req, res, () => {
        db.query('SELECT 1 FROM equipos WHERE id_equipo = ?', [id_equipo], (errEquipo, equipos) => {
            if (errEquipo) {
                console.error('Error al validar el equipo:', errEquipo);
                res.status(500).json({ error: 'Error al actualizar el mantenimiento' });
                return;
            }
            if (equipos.length === 0) {
                res.status(404).json({ error: 'El equipo no existe' });
                return;
            }
            db.query('UPDATE mantenimientos SET id_equipo = ?, tipo = ?, fecha = ?, detalle = ? WHERE id_mantenimiento = ?', [id_equipo, tipo, fecha, detalle, id], (errUpdate) => {
                if (errUpdate) {
                    console.error('Error al actualizar el mantenimiento:', errUpdate);
                    res.status(500).json({ error: 'Error al actualizar el mantenimiento' });
                    return;
                }
                obtenerMantenimiento(id, (errFila, fila) => {
                    if (errFila || !fila) {
                        res.status(500).json({ error: 'El mantenimiento se guardó, pero no se pudo leer el registro' });
                        return;
                    }
                    res.json(fila);
                });
            });
        });
    });
});

// Eliminar un mantenimiento
app.delete('/api/mantenimientos/:id', (req, res) => {
    const { id } = req.params;
    if (!/^\d+$/.test(id)) {
        res.status(400).json({ error: 'Identificador de mantenimiento no válido' });
        return;
    }
    autorizarEdicionMantenimiento(id, req, res, () => {
        db.query('DELETE FROM mantenimientos WHERE id_mantenimiento = ?', [id], (errDelete) => {
            if (errDelete) {
                console.error('Error al eliminar el mantenimiento:', errDelete);
                res.status(500).json({ error: 'Error al eliminar el mantenimiento' });
                return;
            }
            res.json({ mensaje: 'Mantenimiento eliminado correctamente', id_mantenimiento: Number(id) });
        });
    });
});

// Compatibilidad con clientes anteriores que solo corregían la fecha
app.patch('/api/mantenimientos/:id/fecha', (req, res) => {
    const { id } = req.params;
    const fecha = req.body && req.body.fecha;
    const errorFecha = errorDeFecha(fecha);
    if (errorFecha) {
        res.status(400).json({ error: errorFecha });
        return;
    }
    autorizarEdicionMantenimiento(id, req, res, () => {
        db.query('UPDATE mantenimientos SET fecha = ? WHERE id_mantenimiento = ?', [fecha, id], (errUpdate) => {
            if (errUpdate) {
                console.error('Error al actualizar la fecha:', errUpdate);
                res.status(500).json({ error: 'Error al actualizar la fecha' });
                return;
            }
            obtenerMantenimiento(id, (errFila, fila) => {
                if (errFila || !fila) {
                    res.status(500).json({ error: 'La fecha se guardó, pero no se pudo leer el registro' });
                    return;
                }
                res.json(fila);
            });
        });
    });
});

function escaparHtmlServidor(valor) {
    return String(valor ?? '').replace(/[&<>"']/g, (caracter) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[caracter]));
}

async function enviarAlertasPreventivas() {
    const host = process.env.SMTP_HOST;
    const usuarioSMTP = process.env.SMTP_USER;
    const claveSMTP = process.env.SMTP_PASS;
    const remitente = process.env.MAIL_FROM || usuarioSMTP;
    if (!host || !usuarioSMTP || !claveSMTP || !remitente) {
        console.warn('Notificaciones por correo desactivadas: configura SMTP_HOST, SMTP_USER, SMTP_PASS y MAIL_FROM en .env.');
        return 0;
    }

    const [equiposAviso] = await db.promise().query(`
        SELECT e.id_equipo, e.nombre_equipo, e.ubicacion,
               DATE_FORMAT(e.proximo_mto_preventivo, '%Y-%m-%d') AS fecha_programada,
               DATEDIFF(e.proximo_mto_preventivo, CURRENT_DATE) AS dias,
               CASE WHEN e.proximo_mto_preventivo < CURRENT_DATE THEN 'vencido' ELSE 'proximo' END AS estado_alerta
        FROM equipos e
        WHERE e.proximo_mto_preventivo IS NOT NULL
          AND e.proximo_mto_preventivo <= DATE_ADD(CURRENT_DATE, INTERVAL 15 DAY)
        ORDER BY e.proximo_mto_preventivo, e.id_equipo`);
    if (!equiposAviso.length) return 0;

    const ids = [...new Set(equiposAviso.map((equipo) => equipo.id_equipo))];
    const [enviadas] = await db.promise().query(
        'SELECT id_equipo, DATE_FORMAT(fecha_programada, \'%Y-%m-%d\') AS fecha_programada, estado_alerta FROM notificaciones_mantenimiento WHERE id_equipo IN (?)',
        [ids]
    );
    const yaEnviadas = new Set(enviadas.map((fila) => `${fila.id_equipo}|${fila.fecha_programada}|${fila.estado_alerta}`));
    const pendientes = equiposAviso.filter((equipo) => !yaEnviadas.has(`${equipo.id_equipo}|${equipo.fecha_programada}|${equipo.estado_alerta}`));
    if (!pendientes.length) return 0;

    const [usuariosDestino] = await db.promise().query(
        "SELECT DISTINCT correo FROM usuarios WHERE rol IN ('Administrador', 'Tecnico') AND correo IS NOT NULL AND correo <> ''"
    );
    const destinatarios = usuariosDestino.map((usuario) => usuario.correo);
    if (!destinatarios.length) {
        console.warn('No se enviaron alertas: no hay usuarios con correo registrado.');
        return 0;
    }

    const puerto = Number(process.env.SMTP_PORT || 587);
    const transporter = nodemailer.createTransport({
        host,
        port: puerto,
        secure: process.env.SMTP_SECURE ? process.env.SMTP_SECURE === 'true' : puerto === 465,
        auth: { user: usuarioSMTP, pass: claveSMTP }
    });
    const filasHtml = pendientes.map((equipo) => `
        <tr><td>${escaparHtmlServidor(equipo.id_equipo)}</td><td>${escaparHtmlServidor(equipo.nombre_equipo)}</td>
        <td>${escaparHtmlServidor(equipo.ubicacion)}</td><td>${escaparHtmlServidor(equipo.fecha_programada)}</td>
        <td>${equipo.estado_alerta === 'vencido' ? 'Vencido' : `Vence en ${equipo.dias} días`}</td></tr>`).join('');
    await transporter.sendMail({
        from: remitente,
        to: destinatarios.join(', '),
        subject: `BioMant: ${pendientes.length} alerta(s) de mantenimiento preventivo`,
        text: pendientes.map((equipo) => `${equipo.id_equipo} - ${equipo.nombre_equipo}: ${equipo.estado_alerta}, fecha ${equipo.fecha_programada} (${equipo.dias} días).`).join('\n'),
        html: `<p>Equipos con mantenimiento preventivo próximo o vencido:</p><table border="1" cellpadding="6" cellspacing="0"><thead><tr><th>ID</th><th>Equipo</th><th>Ubicación</th><th>Fecha</th><th>Estado</th></tr></thead><tbody>${filasHtml}</tbody></table>`
    });

    const registros = pendientes.map((equipo) => [equipo.id_equipo, equipo.fecha_programada, equipo.estado_alerta]);
    await db.promise().query('INSERT IGNORE INTO notificaciones_mantenimiento (id_equipo, fecha_programada, estado_alerta) VALUES ?', [registros]);
    console.log(`Alertas preventivas enviadas a ${destinatarios.length} usuario(s): ${pendientes.length} equipo(s).`);
    return pendientes.length;
}

const tieneSMTP = Boolean(process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS);
if (tieneSMTP) {
    enviarAlertasPreventivas().catch((error) => console.error('Falló la revisión inicial de alertas preventivas:', error));
    cron.schedule(process.env.MAINTENANCE_ALERT_CRON || '0 8 * * *', () => {
        enviarAlertasPreventivas().catch((error) => console.error('Falló el envío de alertas preventivas:', error));
    }, { timezone: process.env.APP_TIMEZONE || 'America/Bogota' });
    console.log(`Alertas por correo programadas: ${process.env.MAINTENANCE_ALERT_CRON || '0 8 * * *'} (${process.env.APP_TIMEZONE || 'America/Bogota'}).`);
} else {
    console.warn('Notificaciones por correo desactivadas hasta configurar SMTP_HOST, SMTP_USER y SMTP_PASS.');
}

app.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
