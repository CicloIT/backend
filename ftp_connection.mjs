import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import multer from "multer";
import mammoth from "mammoth";
import { google } from 'googleapis';
import cookieParse from 'cookie-parser'
import axios from "axios";
import bcrypt from 'bcrypt';
const { OAuth2 } = google.auth;
import pkg from "exceljs";
const { Workbook } = pkg;
import { Readable } from 'stream';
dotenv.config();
const app = express();
const port = process.env.PORT || 3000;
const CLIENT_ID = process.env.CLIENT_DRIVE;
const CLIENT_SECRET = process.env.CLIENT_SECRET_DRIVE;
app.use(cookieParse());
app.use(cors({
  origin: 'https://lidercom.net.ar', // Remove the trailing slash
  credentials: true,
  methods: ['GET', 'POST', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'Set-Cookie']
}));

app.use(express.json())
/*
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', 'https://lidercom.net.ar'); // Cambia esto por tu dominio real
  res.header('Access-Control-Allow-Credentials', true);
  next();
});*/
// eslint-disable-next-line no-unused-vars
const oAuth2Client = new OAuth2(
  CLIENT_ID,
  CLIENT_SECRET,
  'https://lidercom-7yuo.onrender.com/authentication'
);

async function refreshAccessToken(refreshToken) {
  try {
    const { tokens } = await oAuth2Client.refreshToken(refreshToken);
    return tokens;
  } catch (error) {
    console.error('Error refreshing access token:', error);
    throw new Error('Unable to refresh access token');
  }
}

async function handleAuthenticatedRequest(req, res, next) {
  try {
    oAuth2Client.setCredentials({ access_token: req.cookies.auth_token });
    // Hacer una solicitud a Google Drive
    const drive = google.drive({ version: 'v3', auth: oAuth2Client });
    const response = await drive.files.list();
    // Continuar con la lógica
    next();
  } catch (error) {
    if (error.response && error.response.status === 401) {
      // Token expirado, intentar renovar
      const refreshToken = req.cookies.refresh_token; // Guardaste el refresh token
      if (refreshToken) {
        const newTokens = await refreshAccessToken(refreshToken);
        res.cookie('auth_token', newTokens.access_token, {
          httpOnly: true,
          secure: process.env.NODE_ENV === 'production',
          sameSite: 'lax',
          path: '/',
        });
        // Reintentar la solicitud original
        oAuth2Client.setCredentials({ access_token: newTokens.access_token });
        const drive = google.drive({ version: 'v3', auth: oAuth2Client });
        const response = await drive.files.list();
        next(); // Continúa si tiene éxito
      } else {
        res.status(401).send('Authentication required');
      }
    } else {
      next(error);
    }
  }
}

app.get('/authentication', async (req, res) => {
  const code = req.query.code;
  try {
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    console.log('Tokens en authentication:', tokens);
    console.log("Entro aca");
    res.cookie('auth_token', JSON.stringify(tokens), {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      sameSite: 'none', // Change this to 'none'
      path: '/'
    });
    console.log('Tokens en cookie:', req.cookies.auth_token);
    // Redirect to your frontend
    res.redirect('https://lidercom.net.ar/ftp');
  } catch (error) {
    console.error('Error retrieving access token', error);
    res.status(500).send('Error durante la autenticación');
  }
});

app.get('/checkAuth', (req, res) => {
  console.log('Tokens en cookie:', req.cookies.auth_token);
  const token = req.cookies.auth_token;
  if (token) {
    res.status(200).send('Authenticated');
  } else {
    console.log('Not Authenticated');
    res.status(401).send('Not Authenticated');
  }
});

app.get('/auth', (req, res) => {
  console.log('auth');
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: ['https://www.googleapis.com/auth/drive'],
  });
  res.redirect(authUrl);  // Redirige al usuario a la página de autorización de Google
});

// En tu archivo de rutas
app.get('/listFiles', async (req, res) => {
  const token = req.cookies.auth_token;
  const folderId = req.query.folderId || 'root';
  const searchTerm = req.query.searchTerm || '';
  const mode = req.query.mode || "mydrive";

  if (!token) {
    return res.status(401).send('Unauthorized');
  }

  try {
    const { access_token } = JSON.parse(token);
    oAuth2Client.setCredentials({ access_token });

    const drive = google.drive({ version: 'v3', auth: oAuth2Client });

    let query;
    if (mode === "shared") {
      query = 'sharedWithMe = true'
    }else   
      query = `'${folderId || "root"}' in parents and trashed = false`;
      
    if (searchTerm) {
      query += ` and name contains '${searchTerm}'`;
    }

    const response = await drive.files.list({
      pageSize: 100, // Ajusta el tamaño de la página según tus necesidades
      fields: 'nextPageToken, files(id, name, mimeType, parents)',
      q: query // Filtra los archivos no eliminados
    });

    res.json(response.data.files);
  } catch (error) {
    console.error('Error listing files', error);
    res.status(500).send('Error retrieving files');
  }
});


function getCellAddress(row, column) {
  let columnString = "";
  while (column > 0) {
    column--;
    columnString = String.fromCharCode(65 + (column % 26)) + columnString;
    column = Math.floor(column / 26);
  }
  return columnString + row;
}

app.get("/ExcelFile", async (req, res) => {
  const fileId = req.query.id;
  const token = req.cookies.auth_token;

  if (!fileId || !token) {
    return res.status(400).send('File ID or token missing');
  }

  try {
    const { access_token } = JSON.parse(token);
    oAuth2Client.setCredentials({ access_token });

    const drive = google.drive({ version: 'v3', auth: oAuth2Client });

    // Obtener información del archivo
    const fileMetadata = await drive.files.get({ fileId, fields: 'mimeType' });
    const mimeType = fileMetadata.data.mimeType;

    let response;
    if (mimeType === 'application/vnd.google-apps.spreadsheet') {
      // Exportar Google Sheets como archivo Excel
      response = await drive.files.export(
        { fileId, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
        { responseType: 'arraybuffer' }
      );
    } else {
      // Descargar archivos binarios (como archivos Excel)
      response = await drive.files.get(
        { fileId, alt: 'media' },
        { responseType: 'arraybuffer' }
      );
    }

    try {
      const workbook = new Workbook();
      await workbook.xlsx.load(response.data);
      const worksheet = workbook.worksheets[0];

      const excelData = [];
      const styles = {};
      const mergedCells = {};
      const mergedCellsMap = {};

      Object.entries(worksheet._merges).forEach(([key, mergeRange]) => {
        if (mergeRange && mergeRange.model) {
          const { top, left, bottom, right } = mergeRange.model;
          const startCell = getCellAddress(top, left);
          const endCell = getCellAddress(bottom, right);
          let cellValue;
          try {
            cellValue = worksheet.getCell(startCell).value;
          } catch (error) {
            cellValue = null;
          }

          mergedCells[startCell] = {
            startCell: startCell,
            endCell: endCell,
            value: cellValue,
          };
        } else {
          console.error("  Fusión inválida, falta información del modelo");
        }
      });

      Object.values(mergedCells).forEach(({ startCell, endCell }) => {
        const startRow = parseInt(startCell.replace(/\D/g, ""), 10);
        const startCol = startCell.replace(/\d/g, "");
        const endRow = parseInt(endCell.replace(/\D/g, ""), 10);
        const endCol = endCell.replace(/\d/g, "");

        for (let row = startRow; row <= endRow; row++) {
          for (
            let col = startCol.charCodeAt(0);
            col <= endCol.charCodeAt(0);
            col++
          ) {
            const cellAddress = `${String.fromCharCode(col)}${row}`;
            mergedCellsMap[cellAddress] = startCell;
          }
        }
      });

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const rowData = [];

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const cellAddress = getCellAddress(rowNumber, colNumber);
          let cellValue = cell.value;
          let isMerged = false;

          if (mergedCellsMap[cellAddress]) {
            isMerged = true;
            // Sólo mostrar valor en la celda de inicio de la fusión
            if (cellAddress !== mergedCellsMap[cellAddress]) {
              cellValue = "";
            }
          }

          rowData.push({
            value: cellValue,
            address: cellAddress,
            isMerged: isMerged,
          });

          // Procesamiento de estilos (sin cambios)
          if (cell.style) {
            styles[cellAddress] = {
              fillColor:
                cell.style.fill && cell.style.fill.fgColor
                  ? cell.style.fill.fgColor.argb
                  : undefined,
              fontColor:
                cell.style.font && cell.style.font.color
                  ? cell.style.font.color.argb
                  : undefined,
              border:
                cell.style.border && cell.style.border.top
                  ? cell.style.border.top.color.argb
                  : undefined,
            };
          }
        });

        excelData.push(rowData);
      });
      res.json({
        data: excelData,
        styles: styles,
        mergedCells: mergedCells,
      });

    } catch (error) {
      console.error("Error al leer el archivo con xlsx:", error);
      res.status(500).json({ error: "Error al procesar el archivo Excel" });
    }
  } catch (error) {
    console.error('Error downloading or processing file', error);
    res.status(500).send('Error downloading or processing file: ' + error.message);
  }

});

app.get('/downloadDocument', async (req, res) => {
  const fileId = req.query.id;
  const token = req.cookies.auth_token;

  if (!fileId || !token) {
    return res.status(400).send('File ID or token missing');
  }

  try {
    const { access_token } = JSON.parse(token);
    oAuth2Client.setCredentials({ access_token });

    const drive = google.drive({ version: 'v3', auth: oAuth2Client });

    // Obtener información del archivo
    const fileMetadata = await drive.files.get({ fileId, fields: 'mimeType, name' });
    const { mimeType, name } = fileMetadata.data;

    let content, htmlContent;

    if (mimeType === 'application/vnd.google-apps.document') {
      // Exportar Google Docs como HTML
      const response = await drive.files.export(
        { fileId, mimeType: 'text/html' },
        { responseType: 'text' }
      );
      htmlContent = response.data;
    } else if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      // Descargar archivo .docx y convertir a HTML
      const response = await drive.files.get(
        { fileId, alt: 'media' },
        { responseType: 'arraybuffer' }
      );
      const result = await mammoth.convertToHtml({ buffer: Buffer.from(response.data) });
      htmlContent = result.value;
    } else if (mimeType === 'text/plain') {
      // Descargar archivos de texto y convertir a HTML
      const response = await drive.files.get(
        { fileId, alt: 'media' },
        { responseType: 'text' }
      );
      content = response.data;
      htmlContent = `<pre>${escapeHtml(content)}</pre>`;
    } else {
      // Para otros tipos de archivos, mostrar un mensaje
      htmlContent = '<p>Este tipo de archivo no se puede previsualizar directamente.</p>';
    }

    res.json({
      content: htmlContent,
      mimeType: mimeType,
      name: name
    });

  } catch (error) {
    console.error('Error downloading or processing file', error);
    res.status(500).send('Error downloading or processing file: ' + error.message);
  }
});

// Función auxiliar para escapar HTML
function escapeHtml(unsafe) {
  return unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

app.get('/downloadImage', async (req, res) => {
  const fileId = req.query.id;
  const token = req.cookies.auth_token;

  if (!fileId || !token) {
    return res.status(400).send('File ID or token missing');
  }

  try {
    const { access_token } = JSON.parse(token);
    oAuth2Client.setCredentials({ access_token });

    const drive = google.drive({ version: 'v3', auth: oAuth2Client });

    // Obtener el archivo desde Google Drive
    const response = await drive.files.get(
      { fileId, alt: 'media' },
      { responseType: 'arraybuffer' }
    );

    const mimeType = response.headers['content-type'];
    const buffer = Buffer.from(response.data);

    // Verificar si es una imagen válida
    if (['image/png', 'image/jpeg', 'image/webp'].includes(mimeType)) {
      res.setHeader('Content-Type', mimeType);
      return res.send(buffer);
    } else {
      return res.status(400).send('Invalid image type');
    }
  } catch (error) {
    console.error('Error downloading image', error);
    res.status(500).send('Error downloading image: ' + error.message);
  }
});

app.get('/getThumbnail', async (req, res) => {
  const fileId = req.query.id;
  const token = req.cookies.auth_token;

  if (!fileId) {
    console.error('File ID missing');
    return res.status(400).send('File ID missing');
  }

  if (!token) {
    console.error('Auth token missing');
    return res.status(400).send('Auth token missing');
  }

  try {
    const { access_token } = JSON.parse(token);
    oAuth2Client.setCredentials({ access_token });

    const drive = google.drive({ version: 'v3', auth: oAuth2Client });

    // Obtener la miniatura
    const { data } = await drive.files.get({
      fileId,
      fields: 'thumbnailLink',
    });

    if (data.thumbnailLink) {
      // Descargar la miniatura
      const thumbnailResponse = await axios.get(data.thumbnailLink, { responseType: 'arraybuffer' });
      const mimeType = thumbnailResponse.headers['content-type'];
      const buffer = Buffer.from(thumbnailResponse.data);

      // Establecer el tipo de contenido y enviar la imagen
      res.setHeader('Content-Type', mimeType);
      return res.send(buffer);
    } else {
      console.error('Thumbnail not available');
      return res.status(404).send('Thumbnail not available');
    }
  } catch (error) {
    console.error('Error fetching thumbnail:', error);
    res.status(500).send('Error fetching thumbnail: ' + error.message);
  }
});



const upload = multer({ storage: multer.memoryStorage() });

app.post('/upload', upload.single('file'), async (req, res) => {
  const token = req.cookies.auth_token;
  const folderId = req.body.folderId;
  if (!token) {
    return res.status(401).send('Unauthorized');
  }

  try {
    const { access_token } = JSON.parse(token);
    oAuth2Client.setCredentials({ access_token });

    const drive = google.drive({ version: 'v3', auth: oAuth2Client });
    const file = req.file;
    const fileMetadata = {
      name: file.originalname,
      mimeType: file.mimetype,
      parents: [folderId]
    };
    const media = {
      mimeType: file.mimetype,
      body: Readable.from(file.buffer)
    };

    const response = await drive.files.create({
      resource: fileMetadata,
      media: media,
      fields: 'id'
    });

    res.status(200).send({ fileId: response.data.id });
  } catch (error) {
    console.error('Error uploading file to Google Drive', error);
    res.status(500).send('Error uploading file');
  }
});

app.post('/createFolder', async (req, res) => {
  const token = req.cookies.auth_token;
  const { folderName, parentFolderId } = req.body;

  if (!token) {
    return res.status(401).send('Unauthorized');
  }

  try {
    const { access_token } = JSON.parse(token);
    oAuth2Client.setCredentials({ access_token });

    const drive = google.drive({ version: 'v3', auth: oAuth2Client });

    const folderMetadata = {
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: parentFolderId ? [parentFolderId] : [] // Si no se proporciona, se crea en la raíz
    };

    const response = await drive.files.create({
      resource: folderMetadata,
      fields: 'id, name'
    });

    res.status(200).send({ folderId: response.data.id, folderName: response.data.name });
  } catch (error) {
    console.error('Error creating folder in Google Drive:', error);
    res.status(500).send('Error creating folder');
  }
});


app.delete('/deleteFile', async (req, res) => {
  const fileId = req.query.fileId;
  const token = req.cookies.auth_token;

  if (!fileId || !token) {
    return res.status(400).send({ error: 'No fileId or token provided' });
  }

  try {
    const { access_token } = JSON.parse(token);
    oAuth2Client.setCredentials({ access_token });

    const drive = google.drive({ version: 'v3', auth: oAuth2Client });

    await drive.files.update({
      fileId: fileId,
      requestBody: {
        trashed: true
      }
    });

    res.status(200).send({ message: 'File moved to trash successfully' });
  } catch (error) {
    console.error('Error moving file to trash:', error);
    res.status(500).send({ error: 'Failed to move file to trash' });
  }
});

app.post('/logout', (req, res) => {
  console.log('Intentando borrar cookie:', req.cookies.auth_token);
  res.clearCookie('auth_token', {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'none',
    path: '/',
  });
  console.log('Cookie después de intentar borrar:', req.cookies.auth_token);
  res.status(200).json({ message: 'Logged out successfully' });
});

app.post('/verify-admin-password', async (req, res) => {
  try {
    const { password } = req.body;

    if (!password) {
      return res.status(400).json({ error: 'Password is required' });
    }

    const ADMIN_PASSWORD = '123Liderocom456'; // Asegúrate de almacenar esto de forma segura, como en un archivo de configuración

    const isMatch = await bcrypt.compare(password, ADMIN_PASSWORD);

    if (isMatch) {
      res.json({ success: 'true' });
    } else {
      res.json({ success: 'false' });
    }
  } catch (error) {
    console.error("Error verifying admin password", error);
    res.status(500).json({ error: 'An internal error occurred' });
  }
});

app.get('/downloadPdf', async (req, res) => {
  const fileId = req.query.id;
  const token = req.cookies.auth_token;

  if (!fileId || !token) {
    return res.status(400).send('File ID or token missing');
  }
  try {
    const { access_token } = JSON.parse(token);
    oAuth2Client.setCredentials({ access_token });
    const drive = google.drive({ version: 'v3', auth: oAuth2Client });
    const fileMetadata = await drive.files.get({ fileId, fields: 'mimeType' });
    const mimeType = fileMetadata.data.mimeType;
    if (mimeType !== 'application/pdf') {
      return res.status(400).send('Invalid file type');
    }

    const response = await drive.files.get(
      { fileId, alt: 'media' },
      { responseType: 'arraybuffer' }
    );
    const buffer = Buffer.from(response.data);
    res.setHeader('Content-Type', 'application/pdf');
    res.send(buffer);
  } catch (error) {
    console.error('Error downloading or processing file', error);
    res.status(500).send('Error downloading or processing file: ' + error.message);
  }
})

// Ejemplo de ruta en tu servidor para manejar descargas
app.get('/downloadFile', async (req, res) => {
  const { id } = req.query;
  const token = req.cookies.auth_token;

  if (!token) {
    return res.status(401).send('No authentication token provided');
  }

  try {
    // Parse el token y configura el cliente OAuth
    const parsedToken = JSON.parse(token);
    oAuth2Client.setCredentials(parsedToken);

    // Crear una nueva instancia de drive con la autenticación correcta
    const drive = google.drive({ 
      version: 'v3', 
      auth: oAuth2Client 
    });

    // Obtener los metadatos del archivo primero
    const fileMetadata = await drive.files.get({
      fileId: id,
      fields: 'name, mimeType'
    });

    // Obtener el archivo
    const response = await drive.files.get(
      {
        fileId: id,
        alt: 'media'
      },
      { responseType: 'stream' }
    );

    // Configurar los headers de la respuesta
    res.setHeader('Content-Type', fileMetadata.data.mimeType);
    res.setHeader('Content-Disposition', `attachment; filename="${fileMetadata.data.name}"`);

    // Transmitir el archivo al cliente
    response.data
      .on('end', () => {
        console.log('Download completed');
      })
      .on('error', err => {
        console.error('Error downloading file:', err);
        if (!res.headersSent) {
          res.status(500).send('Error downloading file');
        }
      })
      .pipe(res);

  } catch (error) {
    console.error('Error fetching file from Drive:', error);
    if (!res.headersSent) {
      if (error.code === 401) {
        res.status(401).send('Authentication token expired or invalid');
      } else {
        res.status(500).send(`Error fetching file: ${error.message}`);
      }
    }
  }
});

app.listen(port, '0.0.0.0', () => {
  console.log(`Server is running on port ${port}`);
});