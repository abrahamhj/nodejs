const express = require('express');
const cors = require('cors');
const axios = require('axios');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const https = require('https');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const libre = require('libreoffice-convert');
const memoryCache = new Map();

const app = express();
app.use(cors());
app.use(express.json()); // Para parsear JSON

const CLIENT_ID = "59111103310-q1e7s0qq69uaqn53onplksqss32egr80.apps.googleusercontent.com";
const CLIENT_SECRET = "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";
const REDIRECT_URI = "https://developers.google.com/oauthplayground";
const REFRESH_TOKEN = "1//041r2ycK-7GtDCgYIARAAGAQSNwF-L9Ir9uYcXzUCzXe4CYq91dya0mzZsE9Ys4JQGA86aS7fxI6bmxQWxAMM9NiX9a2RduZBuC0";

const CLOUDMERSIVE_API_KEY = 'e4ba938d-cfb7-42dc-9442-b55233acde41'; // Reemplaza con tu clave API de Cloudmersive

const oauth2Client = new google.auth.OAuth2(
    CLIENT_ID,
    CLIENT_SECRET,
    REDIRECT_URI
);

oauth2Client.setCredentials({ refresh_token: REFRESH_TOKEN });

const drive = google.drive({
    version: 'v3',
    auth: oauth2Client,
});

// Ruta para subir archivos de Word
app.post('/upload', async (req, res) => {
    const { fileName, parametro, folder } = req.body;
    const filePath = path.join(__dirname, 'uploads', fileName);

    if (!fs.existsSync(filePath)) {
        return res.status(404).send('File not found');
    }

    // Modifica el nombre del archivo para incluir el número
    const newFileName = `${parametro}_${fileName}`;

    try {
        const fileMetadata = {
            name: newFileName,
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            parents: [folder]
        };

        const media = {
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            body: fs.createReadStream(filePath),
        };

        const response = await drive.files.create({
            requestBody: fileMetadata,
            media: media,
            fields: 'id'
        });

        res.status(200).json({ message: 'File uploaded successfully', fileId: response.data.id });
    } catch (error) {
        console.error('Failed to upload file:', error.message);
        res.status(500).send('Failed to upload file');
    }
});

// Endpoint para eliminar un archivo
app.delete('/delete/:fileId', async (req, res) => {
    try {
        await drive.files.delete({
            fileId: req.params.fileId,
        });
        res.status(200).send({ message: 'File deleted successfully' });
    } catch (error) {
        console.log(error.message);
        res.status(500).send({ message: 'Failed to delete the file' });
    }
});

// Endpoint para generar una URL pública
app.get('/share/:fileId', async (req, res) => {
    try {
        await drive.permissions.create({
            fileId: req.params.fileId,
            requestBody: {
                role: 'writer',
                type: 'anyone',
            },
        });

        const result = await drive.files.get({
            fileId: req.params.fileId,
            fields: 'webViewLink, webContentLink'
        });

        res.status(200).json({ webViewLink: result.data.webViewLink, webContentLink: result.data.webContentLink });
    } catch (error) {
        console.log(error.message);
        res.status(500).send({ message: 'Failed to create public URL' });
    }
});

app.post('/download', async (req, res) => {
    const { fileId } = req.body;  // Espera recibir el ID del archivo en lugar del URL directo

    try {
        // Primero, obtén el nombre del archivo desde Google Drive
        const fileMetadata = await drive.files.get({
            fileId: fileId,
            fields: 'name'
        });

        const fileName = fileMetadata.data.name;

        // Construye el URL de descarga usando el fileId
        const downloadUrl = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;

        // Ahora inicia la descarga
        const response = await axios({
            method: 'GET',
            url: downloadUrl,
            responseType: 'stream',
            headers: {
                'Authorization': `Bearer ${oauth2Client.credentials.access_token}`
            }
        });

        const filePath = path.join(__dirname, 'downloads', fileName);

        // Crear un write stream para guardar el archivo
        const writer = fs.createWriteStream(filePath);

        response.data.pipe(writer);

        writer.on('finish', () => {
            res.status(200).send({ message: 'File downloaded and saved successfully', path: filePath });
        });

        writer.on('error', (error) => {
            console.error('Error saving the downloaded file:', error.message);
            res.status(500).send({ message: 'Failed to save the downloaded file' });
        });
    } catch (error) {
        console.error('Failed to download file:', error.message);
        res.status(500).send({ message: 'Failed to download file' });
    }
});

app.post('/download-and-send', async (req, res) => {
    const { fileId } = req.body;
    try {
        // Obtiene los metadatos del archivo para establecer el nombre correcto
        const fileMetadata = await drive.files.get({
            fileId: fileId,
            fields: 'name, mimeType'
        });

        const fileName = fileMetadata.data.name;
        const mimeType = fileMetadata.data.mimeType;

        // Descarga el archivo de Google Drive
        const downloadUrl = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
        const response = await axios({
            method: 'GET',
            url: downloadUrl,
            responseType: 'arraybuffer',
            headers: {
                'Authorization': `Bearer ${oauth2Client.credentials.access_token}`
            }
        });

        const downloadsDir = path.join(__dirname, 'downloads');
        if (!fs.existsSync(downloadsDir)) {
            fs.mkdirSync(downloadsDir);
        }

        // Guarda el archivo descargado temporalmente
        const filePath = path.join(__dirname, 'downloads', fileName);
        fs.writeFileSync(filePath, response.data);

        // Convierte el archivo de Word a PDF usando la API de Cloudmersive
        const pdfConversionResponse = await axios({
            method: 'POST',
            url: 'https://api.cloudmersive.com/convert/docx/to/pdf',
            data: fs.readFileSync(filePath),
            headers: {
                'Content-Type': 'application/octet-stream',
                'Apikey': CLOUDMERSIVE_API_KEY
            },
            responseType: 'arraybuffer'
        });

        // Configura los headers de la respuesta para enviar el archivo PDF
        res.setHeader('Content-Disposition', `attachment; filename="${path.basename(fileName, path.extname(fileName))}.pdf"`);
        res.setHeader('Content-Type', 'application/pdf');
        res.send(pdfConversionResponse.data);

        // Elimina el archivo temporal después de enviarlo
        fs.unlink(filePath, (err) => {
            if (err) {
                console.error('Error al eliminar el archivo temporal:', err);
            } else {
                console.log('Archivo temporal eliminado correctamente');
            }
        });
    } catch (error) {
        console.error('Failed to download or send file:', error.message);
        res.status(500).send({ message: 'Failed to download or send file' });
    }
});



async function modifyDocx(inputBuffer, data) {
    const zip = new PizZip(inputBuffer);
    const replacedTags = {};  // Almacena el estado de reemplazo de cada variable

    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        parser: (tag) => {
            return {
                get: () => {
                    if (!replacedTags[tag] && data.hasOwnProperty(tag)) {
                        // Si la variable aún no ha sido reemplazada, reemplázala y márcala como reemplazada
                        replacedTags[tag] = true;
                        return data[tag];
                    } else {
                        // Si ya ha sido reemplazada, devolver el marcador original
                        return `{${tag}}`;
                    }
                }
            };
        }
    });

    doc.render(data);

    // Obtener el buffer del archivo resultante
    let buf = doc.getZip().generate({ type: 'nodebuffer' });

    return buf;
}

async function modifyDocxForPdf(inputBuffer) {
    const zip = new PizZip(inputBuffer);

    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        parser: () => {
            return {
                get: () => ' '  // Reemplaza todas las variables con un espacio en blanco
            };
        }
    });

    doc.render({});  // Render sin datos, todas las variables se reemplazarán por espacios

    // Obtener el buffer del archivo resultante
    let buf = doc.getZip().generate({ type: 'nodebuffer' });

    return buf;
}

async function convertDocxToPdf(inputBuffer) {
    // Crear un nuevo DOCX con los espacios en blanco utilizando la función específica para PDF
    const docWithSpacesBuffer = await modifyDocxForPdf(inputBuffer);

    /*const pdfConversionResponse = await axios({
        method: 'POST',
        url: 'https://api.cloudmersive.com/convert/docx/to/pdf',
        data: docWithSpacesBuffer,
        headers: {
            'Content-Type': 'application/octet-stream',
            'Apikey': CLOUDMERSIVE_API_KEY
        },
        responseType: 'arraybuffer'
    });

    return pdfConversionResponse.data;*/

    return new Promise((resolve, reject) => {
        libre.convert(docWithSpacesBuffer, '.pdf', undefined, (err, done) => {
            if (err) return reject(err);
            resolve(done);
        });
    });
}


// Endpoint para procesar el archivo DOCX y guardarlo en la carpeta 'hojas de ruta'
app.post('/generate-document', async (req, res) => {
    console.time('Tiempo total de procesamiento');
    const uploadDir = path.join(__dirname, 'uploads');
    const outputDir = path.join(__dirname, 'hojas de ruta');

    // Asegurarse de que la carpeta 'hojas de ruta' exista
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir);
    }

    // Leer el código para el nombre del archivo desde el JSON
    const data = req.body;

    // Definir las rutas de los archivos
    const inputDocxPath = path.join(uploadDir, '1 HOJA DE RUTA ELECTRONICA.docx'); // Nombre del archivo DOCX original
    const outputDocxPath = path.join(outputDir, `${data.nrotramite}.docx`); // Nombre del archivo DOCX modificado
    const outputPdfPath = path.join(outputDir, `${data.nrotramite}.pdf`);   // Nombre del archivo PDF generado

    try {
        // Leer el archivo DOCX desde la carpeta 'upload'
        // const inputBuffer = fs.readFileSync(inputDocxPath);
        const inputBuffer = await fs.promises.readFile(inputDocxPath);

        let modifiedDocxBuffer;

        // Verificar si el documento ya existe en la carpeta 'hojas de ruta'
        if (fs.existsSync(outputDocxPath)) {
            // Si existe, leer el documento existente
            const existingDocxBuffer = fs.readFileSync(outputDocxPath);
            // Modificar el documento existente
            modifiedDocxBuffer = await modifyDocx(existingDocxBuffer, data);
        } else {
            // Si no existe, crear uno nuevo
            modifiedDocxBuffer = await modifyDocx(inputBuffer, data);
        }

        // Guardar el DOCX y el PDF
        await fs.promises.writeFile(outputDocxPath, modifiedDocxBuffer);

        res.status(200).send('Documento modificado');
    } catch (err) {
        console.error('Error:', err);
        res.status(500).send('Internal Server Error');
    }
    console.timeEnd('Tiempo total de procesamiento');
});



// Endpoint para recuperar el documento PDF
app.get('/get-document/:nrotramite', async (req, res) => {
    const outputDir = path.join(__dirname, 'hojas de ruta');
    const { nrotramite } = req.params;

    const outputDocxPath = path.join(outputDir, `${nrotramite}.docx`);
    const outputPdfPath = path.join(outputDir, `${nrotramite}.pdf`);

    // Verificar si el archivo DOCX existe
    if (fs.existsSync(outputDocxPath)) {
        try {
            // Variables para almacenar el buffer del PDF
            let pdfBuffer;

            // Verificar si el PDF existe y está actualizado
            let pdfExists = fs.existsSync(outputPdfPath);
            if (pdfExists) {
                const [docxStats, pdfStats] = await Promise.all([
                    fs.promises.stat(outputDocxPath),
                    fs.promises.stat(outputPdfPath)
                ]);

                // Si el DOCX ha sido modificado después del PDF, necesitamos regenerar el PDF
                if (docxStats.mtime > pdfStats.mtime) {
                    pdfExists = false;
                }
            }

            if (pdfExists) {
                // Leer el PDF existente
                pdfBuffer = await fs.promises.readFile(outputPdfPath);
            } else {
                // Leer el DOCX y convertirlo a PDF
                const docxBuffer = await fs.promises.readFile(outputDocxPath);
                pdfBuffer = await convertDocxToPdf(docxBuffer);

                // Guardar el nuevo PDF
                await fs.promises.writeFile(outputPdfPath, pdfBuffer);
            }

            // Enviar el PDF como respuesta
            res.setHeader('Content-Type', 'application/pdf');
            res.send(pdfBuffer);
        } catch (err) {
            console.error('Error al procesar el documento:', err);
            res.status(500).send('Error al obtener el documento');
        }
    } else {
        res.status(404).send('Documento no encontrado');
    }
});


const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
// https.createServer({
//     cert: fs.readFileSync('./certs/server.cer'),
//     key: fs.readFileSync('./certs/server.key')
//   },app).listen(PORT, function(){
//      console.log('Servidor https correindo en el puerto 3000');
//  });