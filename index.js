const express = require('express');
const { google } = require('googleapis');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const stream = require("stream");
const multer = require('multer');
const fs = require('fs');
const { sheetValuesToObject } = require('./utils');
const { config } = require('dotenv');
const cookieParser = require('cookie-parser'); 
config();

const app = express();
const router = express.Router();
const PORT = process.env.PORT || 3001;

let jwtClient = new google.auth.JWT(
  process.env.client_email,
  null,
  (process.env.private_key).replace(/\\n/g, '\n'),
  ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
);

jwtClient.authorize((err) => {
  if (err) {
    console.log('Error connecting to Google APIs:', err);
  } else {
    console.log("Successfully connected to Google APIs!");
  }
});

app.use(bodyParser.json());
app.use(cors());
app.use(cookieParser());

// Middleware para establecer cookies SameSite y secure
app.use((req, res, next) => {
  res.cookie('exampleCookie', 'cookieValue', {
    sameSite: 'None', 
    secure: true, 
    httpOnly: true 
  });
  next();
});

const getSheetRange = (sheetName) => {
  const ranges = {
    'Programas': 'PROGRAMAS!A1:AG1000',
    'Seguimientos': 'SEGUIMIENTOS!A1:H1000',
    'Permisos': 'PERMISOS!A1:C20',
    'Proc_X_Doc': 'PROC_X_PROG_DOCS!A1:E1000',
    'Proc_Fases': 'PROC_FASES!A1:E1000',
    'Proc_X_Prog': 'PROC_X_PROG!A1:C1000',
    'Proc_Fases_Doc': 'PROC_FASES!I1:K1000',
    'Asig_X_Prog': 'ASIG_X_PROG!A1:D1000',
    'Esc_Practica': 'ESC_PRACTICA!A1:D1000',
    'Rel_Esc_Practica': 'REL_ESC_PRACTICA!A1:E1000'
  };
  return ranges[sheetName];
};

const handleSheetRequest = async (req, res, spreadsheetId) => {
  try {
    const { sheetName } = req.body;
    const range = getSheetRange(sheetName);

    if (!range) {
      return res.status(400).json({ error: 'Nombre de hoja no válido' });
    }

    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    res.json({
      status: true,
      data: sheetValuesToObject(response.data.values)
    });
  } catch (error) {
    console.log('Error en la solicitud:', error);
    res.json({
      status: false,
      error: 'Error en la solicitud'
    });
  }
};

router.post('/sendData', async (req, res) => {
  try {
    const { insertData, sheetName } = req.body;
    const spreadsheetId = '1GQY2sWovU3pIBk3NyswSIl_bkCi86xIwCjbMqK_wIAE';
    const range = sheetName;
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    const currentValues = responseSheet.data.values;
    const nextRow = currentValues ? currentValues.length + 1 : 1;
    const updatedRange = `${range}!A${nextRow}`;
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: { values: insertData },
      key: process.env.key,
    });
    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Se insertó correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'No se insertó', status: false });
    }
  } catch (error) {
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

router.post('/', (req, res) => handleSheetRequest(req, res, '1GQY2sWovU3pIBk3NyswSIl_bkCi86xIwCjbMqK_wIAE'));

router.post('/sendDocServ', async (req, res) => {
  try {
    const { insertData, sheetName } = req.body;
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    const range = sheetName;
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    const currentValues = responseSheet.data.values;
    const nextRow = currentValues ? currentValues.length + 1 : 1;
    const updatedRange = `${range}!A${nextRow}`;
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: { values: insertData },
      key: process.env.key,
    });
    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Se insertó correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'No se insertó', status: false });
    }
  } catch (error) {
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

router.post('/docServ', async ( req, res) => {
  
  try {
    const sheets = google.sheets({ version: 'v4',  auth: jwtClient });
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    let range;
    switch (req.body.sheetName) {
      case 'Asig_X_Prog':
        range = 'ASIG_X_PROG!A1:E1000';
        break;
      case 'Esc_Practica':
        range = 'ESC_PRACTICA!A1:D1000';
        break;
      case 'Rel_Esc_Practica':
        range = 'REL_ESC_PRACTICA!A1:G1000';
        break;
      case 'Horario':
        range = 'HORARIOS_PRACT!A1:B1000';
        break;
      case 'firmas':
        range = 'FIRMAS!A1:G1000';
        break;
      default:
        return res.status(400).json({ error: 'Nombre de hoja no válido' });
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key : process.env.key,
    });
    console.log(sheetValuesToObject(response.data.values));   
    res.json({
      status: true, 
      data: sheetValuesToObject(response.data.values)
    }) 
  } catch (error) {
    console.log('error', error); 
    res.json({
      status: false
    })
  }
    
});

router.post('/sendSeguimiento', async (req, res) => {
  try {
    const { insertData, sheetName } = req.body;
    const spreadsheetId = '1BgbiYkp78ylBiiEgjAqPmBze5-aj-GQ081_tFaKw7ys';
    const range = sheetName;
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    const currentValues = responseSheet.data.values;
    const nextRow = currentValues ? currentValues.length + 1 : 1;
    const updatedRange = `${range}!A${nextRow}`;
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: { values: insertData },
      key: process.env.key,
    });
    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Se insertó correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'No se insertó', status: false });
    }
  } catch (error) {
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

router.post('/seguimiento', async ( req, res) => {
  
  try {
    const sheets = google.sheets({ version: 'v4',  auth: jwtClient });
    const spreadsheetId = '1BgbiYkp78ylBiiEgjAqPmBze5-aj-GQ081_tFaKw7ys';
    let range;
    switch (req.body.sheetName) {
      case 'Programas_pm':
        range = 'PROGRAMAS_PM!A1:I1000';
        break;
      case 'Escuela_om':
        range = 'ESCUELAS!A1:AB1000';
        break;
      default:
        return res.status(400).json({ error: 'Nombre de hoja no válido' });
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key : process.env.key,
    });
    console.log(sheetValuesToObject(response.data.values));   
    res.json({
      status: true, 
      data: sheetValuesToObject(response.data.values)
    }) 
  } catch (error) {
    console.log('error', error); 
    res.json({
      status: false
    })
  }
    
});

router.post('/updateSeguimiento', async (req, res) => {
  try {
    const { updateData, id, sheetName } = req.body;
    console.log('Datos recibidos:', req.body);

    if (!updateData || !id || !sheetName) {
      return res.status(400).json({ error: 'Datos faltantes', status: false });
    }

    const spreadsheetId = '1BgbiYkp78ylBiiEgjAqPmBze5-aj-GQ081_tFaKw7ys';
    const range = `${sheetName}!A1:Z1000`;
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    const currentValues = responseSheet.data.values;

    const rowIndex = currentValues.findIndex(row => row[0] == id);
    if (rowIndex === -1) {
      return res.status(404).json({ error: 'ID no encontrado', status: false });
    }

    const updatedRange = `${sheetName}!A${rowIndex + 1}`;
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: { values: [updateData] },
      key: process.env.key,
    });

    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Se actualizó correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'No se actualizó', status: false });
    }
  } catch (error) {
    console.error('Error en la conexión:', error);
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

router.post('/sendReport', async (req, res) => {
  try {
    const { insertData, sheetName } = req.body;
    const spreadsheetId = '1R4Ugfx43AoBjxjsEKYl7qZsAY8AfFNUN_gwcqETwgio';
    const range = sheetName;
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    const currentValues = responseSheet.data.values;
    const nextRow = currentValues ? currentValues.length + 1 : 1;
    const updatedRange = `${range}!A${nextRow}`;
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: { values: insertData },
      key: process.env.key,
    });
    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Se insertó correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'No se insertó', status: false });
    }
  } catch (error) {
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

router.post('/clearSheet', async (req, res) => {
  const { spreadsheetId, sheetName } = req.body;
  try {
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `${sheetName}!A2:Z`,
      key: process.env.key,
    });
    res.status(200).send('Hoja limpiada correctamente');
  } catch (error) {
    console.error('Error al limpiar la hoja:', error);
    res.status(500).send('Error al limpiar la hoja');
  }
});

const upload = multer();

router.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const drive = google.drive({
      version: 'v3',
      auth: jwtClient
    });

    const { originalname } = req.file;

    const fileMetadata = {
      name: originalname,
      parents: ['1Y5dit9wGoK9iKsOd7W-ACjee8zlMhI7R']
    };

    const templateBuffer = Buffer.from(req.file.buffer, 'base64');

    const media = {
      mimeType: req.file.mimetype,
      body: new stream.PassThrough().end(templateBuffer),
    };

    const response = await drive.files.create({
      requestBody: fileMetadata,
      media: media
    });

    console.log('Archivo subido correctamente a Google Drive:', response.data);

    res.json({ success: true, message: 'Archivo subido correctamente a Google Drive', enlace: `https://drive.google.com/file/d/${response.data.id}/view` });
  } catch (error) {
    console.log(error);
    return res.status(400).json({ error: error, status: false });
  }
});

app.use(express.json());
app.use(bodyParser.json());
app.use(cors());
app.use(router);

app.listen(PORT, () => {
  console.log(`Servidor backend escuchando en el puerto ${PORT}`);
});
