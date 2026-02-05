const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const pdf = require('pdf-parse');
const ExcelJS = require('exceljs');
const cors = require('cors');
const https = require('https');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

function downloadPdf(url) {
    return new Promise((resolve, reject) => {
        https.get(url, (res) => {
            if (res.statusCode !== 200) {
                reject(new Error(`Fallo al descargar PDF (Status: ${res.statusCode})`));
                return;
            }
            const data = [];
            res.on('data', (chunk) => data.push(chunk));
            res.on('end', () => resolve(Buffer.concat(data)));
        }).on('error', (err) => reject(err));
    });
}

app.post('/api/process', async (req, res) => {
    const { date } = req.body;
    const [year, month, day] = date.split('-');
    const shortYear = year.substring(2);

    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Transfer-Encoding', 'chunked');

    const sendUpdate = (data) => {
        res.write(JSON.stringify(data) + '\n');
    };

    let browser;
    try {
        const pdfUrl = `https://cibe.bolsadesantiago.com/Documentos/EstadisticasyPublicaciones/Boletines%20Burstiles/ibd${day}${month}${shortYear}.pdf`;

        sendUpdate({ log: `Descargando boletín: ${day}${month}${shortYear}.pdf`, status: 'Descargando', progress: 20 });

        let pdfBuffer;
        try {
            pdfBuffer = await downloadPdf(pdfUrl);
        } catch (e) {
            sendUpdate({ log: 'URL directa fallida. Buscando en la web...' });
            browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
            const page = await browser.newPage();
            await page.goto('https://www.bolsadesantiago.com/estadisticas_boletinbursatil', { waitUntil: 'networkidle2' });
            const searchDate = `${day}-${month}-${year}`;
            const linkHref = await page.evaluate((d) => {
                const a = Array.from(document.querySelectorAll('a.mail-link')).find(x => x.innerText.includes(d));
                return a ? a.href : null;
            }, searchDate);
            if (!linkHref) throw new Error(`Boletín no encontrado para ${searchDate}`);
            pdfBuffer = await downloadPdf(linkHref);
        }

        sendUpdate({ log: 'Extrayendo datos del PDF...', status: 'Procesando', progress: 50 });

        const render_page = (pageData) => {
            return pageData.getTextContent().then(textContent => {
                let lastY, text = '';
                for (let item of textContent.items) {
                    if (lastY == item.transform[5] || !lastY) {
                        text += item.str;
                    } else {
                        text += '\n' + item.str;
                    }
                    lastY = item.transform[5];
                }
                return text + '\n---PAGE_END---';
            });
        };

        const pdfData = await pdf(pdfBuffer, { pagerender: render_page });
        const pages = pdfData.text.split('---PAGE_END---');

        const workbook = new ExcelJS.Workbook();
        let accionesPages = [];
        let cfiPages = [];

        pages.forEach((pageText) => {
            const cleanText = pageText.toLowerCase();
            let sectionType = null;

            if (cleanText.includes('precio de cierre de acciones') || cleanText.includes('precios de cierre de acciones')) {
                sectionType = 'Acciones';
            } else if (
                cleanText.includes('precio de cierre de cfi') ||
                cleanText.includes('precios de cierre de cfi') ||
                cleanText.includes('precio de cierre cfi') ||
                cleanText.includes('precios de cierre cfi') ||
                (cleanText.includes('precios de cierre') && cleanText.includes('mercado cfi'))
            ) {
                sectionType = 'CFI';
            }

            if (sectionType) {
                const tableData = parseActionTable(pageText);
                if (tableData.length > 0) {
                    const formattedRows = tableData.map(row => {
                        let numValue = 0;
                        if (row['Cierre ($)'] && row['Cierre ($)'] !== '0') {
                            const valueStr = row['Cierre ($)'].replace(/\./g, '').replace(',', '.');
                            numValue = parseFloat(valueStr) || 0;
                        }
                        return [row['Nemo'], numValue];
                    });

                    if (sectionType === 'Acciones') {
                        accionesPages.push(formattedRows);
                    } else {
                        cfiPages.push(formattedRows);
                    }
                }
            }
        });

        const createSideBySideTable = (sheetName, allPages, isFirstSheet) => {
            if (allPages.length === 0) return;
            const sheet = workbook.addWorksheet(sheetName);

            let startRow = 1;

            // Add Header Metadata if it's the first sheet
            if (isFirstSheet) {
                const headerInfo = [
                    ['VALORES BURSATILES'],
                    ['Fecha de Proceso:', `${day}/${month}/${year}`],
                    ['Mes:', month],
                    ['Año:', year],
                    ['Archivo Fuente:', `ibd${day}${month}${shortYear}.pdf`],
                    [] // Empty row
                ];
                sheet.addRows(headerInfo);

                // Style Header
                sheet.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF0046AD' } };
                sheet.getCell('A2').font = { bold: true };
                sheet.getCell('A3').font = { bold: true };
                sheet.getCell('A4').font = { bold: true };
                sheet.getCell('A5').font = { bold: true };

                startRow = 7;

            }

            allPages.forEach((rows, pageIdx) => {
                const startCol = pageIdx * 3 + 1;

                sheet.getColumn(startCol).width = 25;
                sheet.getColumn(startCol + 1).width = 15;
                sheet.getColumn(startCol + 2).width = 5;

                const startRef = `${sheet.getColumn(startCol).letter}${startRow}`;

                sheet.addTable({
                    name: `Table_${sheetName.replace(/\s/g, '_')}_${pageIdx}`,
                    ref: startRef,
                    headerRow: true,
                    totalsRow: false,
                    style: {
                        theme: 'TableStyleMedium9',
                        showRowStripes: true,
                    },
                    columns: [
                        { name: 'Nemo', filterButton: true },
                        { name: 'Cierre ($)', filterButton: true },
                    ],
                    rows: rows,
                });

                const priceCol = sheet.getColumn(startCol + 1);
                priceCol.numFmt = '#,##0.000';
            });
        };

        createSideBySideTable('Acciones', accionesPages, true);
        createSideBySideTable('CFI', cfiPages, false);

        if (accionesPages.length === 0 && cfiPages.length === 0) throw new Error('No se encontraron tablas de cierre.');

        const excelFileName = `Boletin_${day}_${month}_${year}.xlsx`;
        const excelPath = path.join(__dirname, 'public', excelFileName);
        await workbook.xlsx.writeFile(excelPath);

        sendUpdate({
            log: 'Proceso completado.',
            status: 'Listo',
            progress: 100,
            fileUrl: `/${excelFileName}`,
            fileName: excelFileName
        });

    } catch (error) {
        sendUpdate({ error: error.message, status: 'Error' });
    } finally {
        if (browser) await browser.close();
        res.end();
    }
});

function parseActionTable(text) {
    const lines = text.split('\n');
    const results = [];
    const fullRowRegex = /^(.*?)((?:[1-9]\d{0,2}(?:\.\d{3})*|0),\d{3})$/;
    const nemoOnlyRegex = /^[A-Z0-9.\-]{3,15}$/;

    lines.forEach(line => {
        const trimmed = line.trim();
        if (!trimmed || trimmed.toLowerCase().includes('nemo') || trimmed.includes('DE 09:00')) return;

        const fullMatch = trimmed.match(fullRowRegex);
        if (fullMatch) {
            const nemo = fullMatch[1].trim();
            const price = fullMatch[2].trim();
            if (nemo && !nemo.includes('Nemotécnico')) {
                results.push({ 'Nemo': nemo, 'Cierre ($)': price });
            }
        } else if (trimmed.match(nemoOnlyRegex)) {
            results.push({ 'Nemo': trimmed, 'Cierre ($)': '0' });
        }
    });
    return results;
}

app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
