const fs = require('node:fs');
const { cwd } = require('node:process');
const path = require('node:path');
const { PDFDocument, rgb } = require('pdf-lib');
const xlsx = require('xlsx');
const fontkit = require('fontkit');
const colors = require('colors');

const srcPath = path.join(cwd(), 'src');
const outputDir = path.join(srcPath, 'certificadosGDP');

// Verificar si el directorio existe, si no, crearlo
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// Leer el archivo Excel
const workbook = xlsx.readFile(path.join(srcPath, 'CONTROL.xlsx'));
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const datos = xlsx.utils.sheet_to_json(sheet);

// Función para centrar el texto
function drawCenteredText(page, text, { x, y, size, font, color, maxWidth }) {
    const textWidth = font.widthOfTextAtSize(text, size);
    const textX = x + (maxWidth - textWidth) / 2;
    page.drawText(text, {
        x: textX,
        y: y,
        size: size,
        font: font,
        color: color,
        lineHeight: 18,
        maxWidth: maxWidth,
    });
}

// Función para dividir el texto en varias líneas
function splitTextIntoLines(text, font, size, maxWidth) {
    const words = text.split(' ');
    const lines = [];
    let currentLine = '';

    for (const word of words) {
        const testLine = currentLine + (currentLine ? ' ' : '') + word;
        if (font.widthOfTextAtSize(testLine, size) > maxWidth) {
            lines.push(currentLine);
            currentLine = word;
        } else {
            currentLine = testLine;
        }
    }
    lines.push(currentLine);
    return lines;
}

// Función para crear un certificado
async function crearCertificado(data) {
    const pdfDoc = await PDFDocument.create();
    
    // Registrar fontkit
    pdfDoc.registerFontkit(fontkit);

    const page = pdfDoc.addPage([841.89, 595.28]); // A4 horizontal

    // Cargar las fuentes
    const calibriFontLightBytes = fs.readFileSync(path.join(srcPath, 'fonts', 'calibri.ttf'));
    const calibriLightFont = await pdfDoc.embedFont(calibriFontLightBytes);

    const calibriFontBoldBytes = fs.readFileSync(path.join(srcPath, 'fonts', 'calibrib.ttf'));
    const calibriBoldFont = await pdfDoc.embedFont(calibriFontBoldBytes);

    // Cargar imagen de fondo
    const backgroundImageBytes = fs.readFileSync(path.join(srcPath, 'img', 'certificado.png'));
    const backgroundImage = await pdfDoc.embedPng(backgroundImageBytes);

    // Dibujar la imagen de fondo
    page.drawImage(backgroundImage, {
        x: 0,
        y: 0,
        width: 841.89,
        height: 595.28,
    });

    // Añadir texto al PDF
    const { width, height } = page.getSize();
    const fontSize = 12;
    const headerFontSize = 24;

    const postionX = 60;

    drawCenteredText(page, `${data['Nombres y Apellidos']}`, {
        x: postionX,
        y: height - 280,
        size: headerFontSize,
        font: calibriBoldFont,
        color: rgb(0, 0, 0),
        maxWidth: width - 100,
    });

    const lines1 = splitTextIntoLines('Por haber participado como ASISTENTE en el curso de:', calibriLightFont, fontSize, width - 100);
    let yOffset = height - 300;
    lines1.forEach(line => {
        drawCenteredText(page, line, {
            x: postionX,
            y: yOffset,
            size: fontSize,
            font: calibriLightFont,
            color: rgb(0, 0, 0),
            maxWidth: width - 100,
        });
        yOffset -= 20;
    });

    const lines2 = splitTextIntoLines(`${data['Mención del Certificado']}`, calibriBoldFont, fontSize, width - 100);
    lines2.forEach(line => {
        drawCenteredText(page, line, {
            x: postionX - 5,
            y: yOffset,
            size: fontSize,
            font: calibriBoldFont,
            color: rgb(0, 0, 0),
            maxWidth: width - 100,
        });
        yOffset -= 20;
    });

    const lines3 = splitTextIntoLines(`desarrollado del ${data['FECHA DE INICIO']} al ${data['FECHA DE TERMINO']}`, calibriLightFont, fontSize, width - 100);
    lines3.forEach(line => {
        drawCenteredText(page, line, {
            x: postionX,
            y: yOffset,
            size: fontSize,
            font: calibriLightFont,
            color: rgb(0, 0, 0),
            maxWidth: width - 100,
        });
        yOffset -= 20;
    });

    const lines4 = splitTextIntoLines(`en la modalidad ${data['MODALIDAD'] || 'virtual'} y con una duración de ${data['HORAS']} horas pedagógicas.`, calibriLightFont, fontSize, width - 100);
    lines4.forEach(line => {
        drawCenteredText(page, line, {
            x: postionX,
            y: yOffset,
            size: fontSize,
            font: calibriLightFont,
            color: rgb(0, 0, 0),
            maxWidth: width - 100,
        });
        yOffset -= 20;
    });

    const lines5 = splitTextIntoLines(`${data['CIUDAD'] || 'Chepén'}; ${data['FECHA DE EMISION']}`, calibriLightFont, fontSize, width - 100);
    lines5.forEach(line => {
        drawCenteredText(page, line, {
            x: postionX + 280,
            y: yOffset,
            size: fontSize,
            font: calibriLightFont,
            color: rgb(0, 0, 0),
            maxWidth: width - 100,
        });
        yOffset -= 20;
    });

    const lines6 = splitTextIntoLines(`${data['N° de Registro']}`, calibriLightFont, 8, width - 100);
    lines6.forEach(line => {
        drawCenteredText(page, line, {
            x: postionX + 290,
            y: yOffset - 74,
            size: 8,
            font: calibriLightFont,
            color: rgb(0, 0, 0),
            maxWidth: width - 100,
        });
        yOffset -= 20;
    });

    const lines7 = splitTextIntoLines(`${data['N° de Código']}`, calibriLightFont, 8, width - 100);
    lines7.forEach(line => {
        drawCenteredText(page, line, {
            x: postionX + 270,
            y: yOffset - 74,
            size: 8,
            font: calibriLightFont,
            color: rgb(0, 0, 0),
            maxWidth: width - 100,
        });
        yOffset -= 20;
    });

    const lines8 = splitTextIntoLines(`${data['N° Libro']}`, calibriLightFont, 8, width - 100);
    lines8.forEach(line => {
        drawCenteredText(page, line, {
            x: postionX + 230,
            y: yOffset - 72,
            size: 8,
            font: calibriLightFont,
            color: rgb(0, 0, 0),
            maxWidth: width - 100,
        });
        yOffset -= 20;
    });

    // Guardar el PDF con el nombre deseado en la ruta correcta
    const pdfBytes = await pdfDoc.save();
    const nombrePdf = `${data['N° Libro']}_${data['N° de Código']}.pdf`;
    fs.writeFileSync(path.join(outputDir, nombrePdf), pdfBytes);
}

// Crear certificados y mostrar el progreso general
async function procesarCertificados(datos, batchSize = 5) {
    const totalCertificados = datos.length;
    let certificadosProcesados = 0;

    for (let i = 0; i < datos.length; i += batchSize) {
        const batch = datos.slice(i, i + batchSize);
        await Promise.all(batch.map(async (data) => {
            await crearCertificado(data);
        }));

        certificadosProcesados += batch.length;
        const progreso = ((certificadosProcesados / totalCertificados) * 100).toFixed(2);
        console.log(`Generando Certificados: ${colors.green(progreso,'%')} \u{23F1}`);
    }

    console.log('Todos los certificados se han generado correctamente.');
}

// Llamar a la función con los datos y el tamaño del lote deseado
procesarCertificados(datos);
