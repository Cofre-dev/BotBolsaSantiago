const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const pdf = require('pdf-parse');
const XLSX = require('xlsx');
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

        const workbook = XLSX.utils.book_new();
        let sheetCount = 0;

        pages.forEach((pageText) => {
            const cleanText = pageText.toLowerCase();
            if (cleanText.includes('precio de cierre de acciones') || cleanText.includes('precios de cierre de acciones')) {
                const tableData = parseActionTable(pageText);
                if (tableData.length > 3) {
                    sheetCount++;

                    // Convert rows to numeric values properly for Excel
                    const numericData = tableData.map(row => {
                        // row['Cierre ($)'] comes like "1.101,400"
                        // We remove dots (thousands) and replace comma with dot (decimal)
                        // Then parse it as a number so XLSX knows it's numeric.
                        const valueStr = row['Cierre ($)'].replace(/\./g, '').replace(',', '.');
                        const numValue = parseFloat(valueStr);

                        return {
                            'Nemo': row['Nemo'],
                            'Cierre ($)': numValue
                        };
                    });

                    const worksheet = XLSX.utils.json_to_sheet(numericData);
                    XLSX.utils.book_append_sheet(workbook, worksheet, `hoja${sheetCount}`);
                }
            }
        });

        if (sheetCount === 0) throw new Error('No se encontraron tablas de cierre de acciones.');

        const excelFileName = `Boletin_${day}_${month}_${year}.xlsx`;
        const excelPath = path.join(__dirname, 'public', excelFileName);
        XLSX.writeFile(workbook, excelPath);

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
    const lineRegex = /^(.*?)\s?((?:[1-9]\d{0,2}(?:\.\d{3})*|0),\d{3})$/;

    lines.forEach(line => {
        const trimmed = line.trim();
        if (!trimmed || trimmed.toLowerCase().includes('nemo') || trimmed.includes('DE 09:00')) return;

        const match = trimmed.match(lineRegex);
        if (match) {
            const nemo = match[1].trim();
            const price = match[2].trim();
            if (nemo && nemo.length > 2 && isNaN(nemo)) {
                results.push({ 'Nemo': nemo, 'Cierre ($)': price });
            }
        }
    });
    return results;
}

app.listen(PORT);
