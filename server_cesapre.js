/**
 * CESAPRE - Backend Node.js
 * Procesa pagos Culqi y genera sesiones MINEDU automáticamente
 */

const express = require('express');
const axios = require('axios');
const cors = require('cors');
const path = require('path');
const { Document, Packer, Table, TableRow, TableCell, Paragraph, TextRun, AlignmentType, BorderStyle } = require('docx');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('.'));

// Configuración Culqi
const CULQI_SECRET_KEY = process.env.CULQI_SECRET_KEY || 'sk_test_7Bq0aN2mkOoWKesQ';
const CULQI_API_URL = 'https://api.culqi.com/v2/charges';

// PALETA DE COLORES MINEDU (7 colores)
const COLORES = {
    datos: { hex: '#003366', rgb: '0, 51, 102' },           // Azul marino
    area: { hex: '#1B5E20', rgb: '27, 94, 32' },            // Verde oscuro
    competencias: { hex: '#CC5200', rgb: '204, 82, 0' },    // Naranja
    criterios: { hex: '#A30052', rgb: '163, 0, 82' },       // Rojo oscuro
    enfoques: { hex: '#667eea', rgb: '102, 126, 234' },     // Morado
    materiales: { hex: '#FFD700', rgb: '255, 215, 0' },     // Amarillo
    momentos: { hex: '#8B4513', rgb: '139, 69, 19' },       // Marrón
    borde: '#CCCCCC'                                         // Gris claro para bordes
};

// ==================== RUTAS ====================

/**
 * POST /api/procesar-pago
 * Recibe token de Culqi y procesa pago
 */
app.post('/api/procesar-pago', async (req, res) => {
    try {
        const { token, email, tema, edad, area, duracion, docente, institucion } = req.body;

        if (!token || !email) {
            return res.status(400).json({ 
                success: false, 
                error: 'Token y email requeridos' 
            });
        }

        // Procesar pago con Culqi
        const charge = await procesarCulqi(token, email);

        if (!charge.success) {
            return res.status(400).json({ 
                success: false, 
                error: 'Error en el pago' 
            });
        }

        // Generar cesión
        const docxBuffer = await generarCesion({
            tema, edad, area, duracion, docente, institucion
        });

        // Responder con buffer
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': `attachment; filename="CESION_${tema}_${edad}anos.docx"`
        });

        res.send(docxBuffer);

    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ 
            success: false, 
            error: error.message 
        });
    }
});

/**
 * Procesar pago con Culqi
 */
async function procesarCulqi(token, email) {
    try {
        const response = await axios.post(CULQI_API_URL, {
            amount: 100, // S/ 1.00 en centimos
            currency_code: 'PEN',
            email: email,
            source_id: token
        }, {
            headers: {
                'Authorization': `Bearer ${CULQI_SECRET_KEY}`,
                'Content-Type': 'application/json'
            }
        });

        return {
            success: response.data.object === 'charge' && response.data.status === 'succeeded',
            data: response.data
        };

    } catch (error) {
        console.error('Error Culqi:', error.message);
        return { success: false };
    }
}

/**
 * Generar sesión MINEDU en Word
 */
async function generarCesion(datos) {
    const { tema, edad, area, duracion, docente, institucion } = datos;

    // Mapeos
    const temaMap = {
        'quien-soy': 'Quién soy',
        'mi-cuerpo': 'Mi cuerpo',
        'mi-familia': 'Mi familia',
        'emociones': 'Emociones'
    };

    const areaMap = {
        'ps': 'Personal Social',
        'comunicacion': 'Comunicación',
        'matematica': 'Matemática',
        'ct': 'Ciencia y Tecnología',
        'psicomotriz': 'Psicomotriz'
    };

    const temaFormato = temaMap[tema] || tema;
    const areaFormato = areaMap[area] || area;

    // Crear documento
    const doc = new Document({
        sections: [{
            children: [
                // TÍTULO
                new Paragraph({
                    text: 'SESIÓN DE APRENDIZAJE',
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 400 },
                    run: { bold: true, size: 32 }
                }),

                // TABLA 1: DATOS INFORMATIVOS
                crearTablaInformativa(temaFormato, areaFormato, edad, duracion, docente, institucion),

                new Paragraph({ text: '', spacing: { after: 200 } }),

                // TABLA 2: LISTA DE COTEJO
                crearTablaListaCotejo(),

                new Paragraph({ text: '', spacing: { after: 200 } }),

                // TABLA 3: SIMILITUDES Y DIFERENCIAS
                crearTablaSimilitudes(),

                new Paragraph({ text: '', spacing: { after: 400 } }),

                // REFLEXIONES
                new Paragraph({
                    text: 'REFLEXIONES DEL DOCENTE',
                    spacing: { before: 200, after: 200 },
                    run: { bold: true, size: 24, color: COLORES.competencias.hex.slice(1) }
                }),

                new Paragraph('¿Qué fortalezas observé en mis estudiantes durante la sesión?'),
                new Paragraph(''),
                new Paragraph('¿Qué dificultades presentaron y cómo las abordé?'),
                new Paragraph(''),
                new Paragraph('¿Qué ajustes realizaré en la siguiente sesión?'),
                new Paragraph(''),
                new Paragraph('¿Cómo evidencié el logro de los desempeños esperados?'),

                new Paragraph({ text: '', spacing: { after: 400 } }),

                // REFERENCIAS
                new Paragraph({
                    text: 'REFERENCIAS BIBLIOGRÁFICAS',
                    spacing: { before: 200, after: 200 },
                    run: { bold: true, size: 24, color: COLORES.competencias.hex.slice(1) }
                }),

                new Paragraph('MINEDU (2016). Programa Curricular de Educación Inicial. https://www.gob.pe/minedu'),
                new Paragraph('MINEDU (2016). Orientaciones para el Desarrollo Curricular. https://www.minedu.gob.pe'),
                new Paragraph('Flores Ochoa, R. (2000). Evaluación, Aprendizaje y Competencias. Editorial Magisterio.')
            ]
        }]
    });

    // Generar buffer
    const buffer = await Packer.toBuffer(doc);
    return buffer;
}

/**
 * Crear tabla informativa (TABLA 1)
 */
function crearTablaInformativa(tema, area, edad, duracion, docente, institucion) {
    const cells = [
        crearCeldaEncabezado('DATOS INFORMATIVOS', COLORES.datos.hex),
        new TableCell({ children: [new Paragraph(institucion || 'Institución Educativa')] }),
        crearCeldaEncabezado('GRADO', COLORES.datos.hex),
        new TableCell({ children: [new Paragraph(`${edad} años`)] }),
        crearCeldaEncabezado('DOCENTE', COLORES.datos.hex),
        new TableCell({ children: [new Paragraph(docente || 'Nombre del Docente')] }),
        crearCeldaEncabezado('FECHA', COLORES.datos.hex),
        new TableCell({ children: [new Paragraph(new Date().toLocaleDateString('es-PE'))] }),
        crearCeldaEncabezado('ÁREA CURRICULAR', COLORES.area.hex),
        new TableCell({ children: [new Paragraph(area)] }),
        crearCeldaEncabezado('COMPETENCIA', COLORES.competencias.hex),
        new TableCell({ children: [new Paragraph('Competencia del área')] }),
        crearCeldaEncabezado('DESEMPEÑOS', COLORES.competencias.hex),
        new TableCell({ children: [new Paragraph(`Desempeños esperados para ${edad} años`)] }),
        crearCeldaEncabezado('CRITERIOS DE EVALUACIÓN', COLORES.criterios.hex),
        new TableCell({ children: [new Paragraph('Criterio 1\nCriterio 2\nCriterio 3')] }),
    ];

    return new Table({
        rows: [
            new TableRow({ children: [cells[0], cells[1]] }),
            new TableRow({ children: [cells[2], cells[3]] }),
            new TableRow({ children: [cells[4], cells[5]] }),
            new TableRow({ children: [cells[6], cells[7]] }),
            new TableRow({ children: [cells[8], cells[9]] }),
            new TableRow({ children: [cells[10], cells[11]] }),
            new TableRow({ children: [cells[12], cells[13]] }),
            new TableRow({ children: [cells[14], cells[15]] }),
        ],
        width: { size: 100, type: 'pct' },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            left: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            right: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            insideVertical: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde }
        }
    });
}

/**
 * Crear tabla lista de cotejo (TABLA 2)
 */
function crearTablaListaCotejo() {
    const headerCells = [
        crearCeldaEncabezado('N°', COLORES.competencias.hex),
        crearCeldaEncabezado('Apellidos y Nombres', COLORES.competencias.hex),
        crearCeldaEncabezado('Criterio 1', COLORES.competencias.hex),
        crearCeldaEncabezado('Criterio 2', COLORES.competencias.hex),
        crearCeldaEncabezado('Mejora', COLORES.competencias.hex)
    ];

    const rows = [new TableRow({ children: headerCells })];

    // Agregar 10 filas de estudiantes
    for (let i = 1; i <= 10; i++) {
        rows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph(i.toString())] }),
                new TableCell({ children: [new Paragraph('')] }),
                new TableCell({ children: [new Paragraph('')] }),
                new TableCell({ children: [new Paragraph('')] }),
                new TableCell({ children: [new Paragraph('')] })
            ]
        }));
    }

    return new Table({
        rows: rows,
        width: { size: 100, type: 'pct' },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            left: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            right: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            insideVertical: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde }
        }
    });
}

/**
 * Crear tabla similitudes y diferencias (TABLA 3)
 */
function crearTablaSimilitudes() {
    return new Table({
        rows: [
            new TableRow({
                children: [
                    crearCeldaEncabezado('SIMILITUDES', COLORES.competencias.hex),
                    crearCeldaEncabezado('DIFERENCIAS', COLORES.competencias.hex)
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph('')] }),
                    new TableCell({ children: [new Paragraph('')] })
                ]
            })
        ],
        width: { size: 100, type: 'pct' },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            left: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            right: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde },
            insideVertical: { style: BorderStyle.SINGLE, size: 6, color: COLORES.borde }
        }
    });
}

/**
 * Crear celda de encabezado con color
 */
function crearCeldaEncabezado(texto, color) {
    return new TableCell({
        children: [new Paragraph({
            text: texto,
            run: { bold: true, color: 'FFFFFF', size: 20 }
        })],
        shading: { fill: color.slice(1) }
    });
}

// ==================== INICIAR SERVIDOR ====================

app.listen(PORT, () => {
    console.log(`✅ CESAPRE Backend corriendo en puerto ${PORT}`);
    console.log(`📧 Configurar CULQI_SECRET_KEY: export CULQI_SECRET_KEY=sk_test_xxxxx`);
});
