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

// Configuración de multer para almacenamiento en memoria
const upload = multer({ storage: multer.memoryStorage() });

// ID de la carpeta principal en Google Drive
const parentFolderId = '178MjHfhOhkZCs4mOAtXdJ8EjDW08PEoC';

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

// Inicializar el cliente de la API de Google Drive
const drive = google.drive({ version: 'v3', auth: jwtClient });

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

//Función para obtener el rango de las hojas de cálculo 
const getSheetRange = (sheetName) => {
  const ranges = {
    'Programas': 'PROGRAMAS!A1:AH1000',
    'Seguimientos': 'SEGUIMIENTOS!A1:H2000',
    'Permisos': 'PERMISOS!A1:C100',
    'Proc_X_Doc': 'PROC_X_PROG_DOCS!A1:E1000',
    'Proc_Fases': 'PROC_FASES!A1:E1000',
    'Proc_X_Prog': 'PROC_X_PROG!A1:C1000',
    'Proc_Fases_Doc': 'PROC_FASES!I1:K1000',
    'Asig_X_Prog': 'ASIG_X_PROG!A1:D1000',
    'Esc_Practica': 'ESC_PRACTICA!A1:D1000',
    'Rel_Esc_Practica': 'REL_ESC_PRACTICA!A1:E1000',
    'HISTORICO': 'HISTORICO!A1:G1000',
    'ESTADISTICAS': 'ESTADISTICAS!A1:Q2000'
  };
  return ranges[sheetName];
};

//Función para manejar las solicitudes de hojas de cálculo
const handleSheetRequest = async (req, res, spreadsheetId) => {
  try {
    const { sheetName } = req.body;
    const range = getSheetRange(sheetName);

    console.log(`Solicitud para hoja: ${sheetName}, Range: ${range}`);

    if (!range) {
      console.log('Hojas disponibles:', Object.keys({
        'Programas': 'PROGRAMAS!A1:AH1000',
        'Seguimientos': 'SEGUIMIENTOS!A1:H2000',
        'Permisos': 'PERMISOS!A1:C100',
        'Proc_X_Doc': 'PROC_X_PROG_DOCS!A1:E1000',
        'Proc_Fases': 'PROC_FASES!A1:E1000',
        'Proc_X_Prog': 'PROC_X_PROG!A1:C1000',
        'Proc_Fases_Doc': 'PROC_FASES!I1:K1000',
        'Asig_X_Prog': 'ASIG_X_PROG!A1:D1000',
        'Esc_Practica': 'ESC_PRACTICA!A1:D1000',
        'Rel_Esc_Practica': 'REL_ESC_PRACTICA!A1:E1000',
        'HISTORICO': 'HISTORICO!A1:G1000',
        'ESTADISTICAS': 'ESTADISTICAS!A1:Q2000'
      }));
      return res.status(400).json({ error: `Nombre de hoja no válido: ${sheetName}` });
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
    console.log(`Error en la solicitud para ${req.body.sheetName}:`, error);
    res.json({
      status: false,
      error: `Error en la solicitud para ${req.body.sheetName}: ${error.message}`
    });
  }
};

//Función para enviar datos a las hojas de cálculo
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

//Función para obtener datos de las hojas de cálculo principales
router.post('/', (req, res) => handleSheetRequest(req, res, '1GQY2sWovU3pIBk3NyswSIl_bkCi86xIwCjbMqK_wIAE'));

//Función para enviar datos a la hoja de cálculo de Docencia y Servicio
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

//Función para obtener datos de la hoja de cálculo de Docencia y Servicio
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
      case 'anexos':
        range = 'ANEXOS_TEC!A1:G1000';
        break;
      case 'Programas':
        range = 'PROGRAMAS!A1:AH1000';
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

//Función para enviar datos a la hoja de cálculo de Seguimiento PM
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

//Función para obtener datos de la hoja de cálculo de Seguimiento PM
router.post('/seguimiento', async ( req, res) => {
  
  try {
    const sheets = google.sheets({ version: 'v4',  auth: jwtClient });
    const spreadsheetId = '1BgbiYkp78ylBiiEgjAqPmBze5-aj-GQ081_tFaKw7ys';
    let range;
    switch (req.body.sheetName) {
      case 'Programas_pm':
        range = 'PROGRAMAS_PM!A1:AZ1000';
        break;
      case 'Escuela_om':
        range = 'ESCUELAS!A1:AZ1000';
        break;
      case 'Programas':
        range = 'PROGRAMAS!A1:AZ1000';
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

//Función para actualizar los datos de la hoja de cálculo de Seguimiento PM
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

//Función para enviar los datos a la hoja de cálculo Reportes Actividades
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

//Función para limpiar los datos de la hoja de cálculo Reportes Actividades
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

// Endpoint para subir archivos a Google Drive
router.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: 'No se subió ningún archivo' });
    }

    const scenarioName = req.body.scenarioName || 'Archivos Generales';
    
    // 1. Buscar si la carpeta del escenario ya existe
    const folderQuery = `mimeType='application/vnd.google-apps.folder' and name='${scenarioName}' and '${parentFolderId}' in parents and trashed=false`;
    let folderResponse = await drive.files.list({
        q: folderQuery,
        fields: 'files(id, name)',
        spaces: 'drive'
    });

    let folderId;
    if (folderResponse.data.files.length > 0) {
      // La carpeta ya existe
      folderId = folderResponse.data.files[0].id;
    } else {
      // La carpeta no existe, hay que crearla
      const folderMetadata = {
        name: scenarioName,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [parentFolderId]
      };
      const newFolder = await drive.files.create({
        resource: folderMetadata,
        fields: 'id'
      });
      folderId = newFolder.data.id;
    }

    // 2. Subir el archivo a la carpeta del escenario
    const fileMetadata = {
      name: req.file.originalname,
      parents: [folderId]
    };
    const media = {
      mimeType: req.file.mimetype,
      body: stream.Readable.from(req.file.buffer)
    };

    const uploadedFile = await drive.files.create({
      resource: fileMetadata,
      media: media,
      fields: 'id, webViewLink'
    });

    // 3. Hacer el archivo públicamente visible (opcional pero recomendado para URLs)
    await drive.permissions.create({
      fileId: uploadedFile.data.id,
      resource: {
        role: 'reader',
        type: 'anyone'
      }
    });

    // 4. Devolver la URL del archivo
    res.json({
      success: true,
      message: 'Archivo subido correctamente a Google Drive',
      fileUrl: uploadedFile.data.webViewLink
    });

  } catch (error) {
    console.error('Error al subir el archivo a Google Drive:', error);
    res.status(500).json({ success: false, message: 'Error interno al subir el archivo' });
  }
});

//Función para editar los detalles de los programas 
app.post('/updateData', async (req, res) => {
  try {
    const { id, Sede, Facultad, Escuela, Departamento, Sección, 'Nivel de Formación': NivelFormacion, 'Titulo a Conceder': Titulo, Jornada, Modalidad, Créditos, Periodicidad, Duración, 'Fecha RRC': FechaRRC, 'Fecha RAAC': FechaRAAC, Acreditable, Contingencia} = req.body;
    if (!id) {
      return res.status(400).json({ error: 'ID faltante', status: false });
    }

    const spreadsheetId = '1GQY2sWovU3pIBk3NyswSIl_bkCi86xIwCjbMqK_wIAE';
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const range = 'PROGRAMAS!A1:AE1000';  
    const response = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    const rows = response.data.values;

    const rowIndex = rows.findIndex(row => row[29] == id); 
    if (rowIndex === -1) {
      return res.status(404).json({ error: 'ID no encontrado', status: false });
    }

    const updatedRow = [
      rows[rowIndex][0], // Programa Académico
      rows[rowIndex][1], // Snies
      Sede,
      Facultad,
      Escuela,
      Departamento,
      Sección,
      NivelFormacion, //posgrado/pregrado
      rows[rowIndex][8], // nivel de formacion
      Titulo,
      Jornada,
      Modalidad,
      Créditos,
      Periodicidad,
      Duración,
      rows[rowIndex][15], // Cupos
      Acreditable, // Acreditable
      rows[rowIndex][17], // Estado
      rows[rowIndex][18], // Tipo de Creación
      rows[rowIndex][19], // RC Vigente
      rows[rowIndex][20], // FechaExpedRC
      rows[rowIndex][21], // DuracionRC
      FechaRRC,
      rows[rowIndex][23], // Fase RRC
      rows[rowIndex][24], // AC Vigente
      rows[rowIndex][25], // FechaExpedAC
      rows[rowIndex][26], // DuracionAC
      FechaRAAC,
      rows[rowIndex][28], // FASE RAC
      id,
      rows[rowIndex][30], // AAC_1A
      rows[rowIndex][31], // MOD
      rows[rowIndex][32], // MOD_SUS
      Contingencia, 
    ];

    const updateRange = `PROGRAMAS!A${rowIndex + 1}:AH${rowIndex + 1}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updateRange,
      valueInputOption: 'RAW',
      resource: { values: [updatedRow] },
    });

    res.status(200).json({ success: 'Datos actualizados correctamente', status: true });
  } catch (error) {
    console.error('Error al actualizar datos:', error);
    res.status(500).json({ error: 'Error al actualizar datos', details: error.message, status: false });
  }
});

//Función para obtener los anexos de la hoja de calculo de Docencia y Servicio
router.post('/getAnexos', async (req, res) => {
  try {
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    const range = 'ANEXOS_TEC!A1:M1000';

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });

    res.json({
      status: true,
      data: sheetValuesToObject(response.data.values) // Convierte los valores de la hoja en objetos
    });
  } catch (error) {
    console.log('Error en la solicitud:', error);
    res.json({
      status: false,
      error: 'Error en la solicitud'
    });
  }
});

//Función para actualizar los anexos de la hoja de calculo de Docencia y Servicio
router.post('/updateAnexo', async (req, res) => {
  try {
    const { updateData, id } = req.body;
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    const range = 'ANEXOS_TEC!A1:M1000';
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    const currentValues = responseSheet.data.values;

    // Encontrar la fila con el ID especificado
    const rowIndex = currentValues.findIndex(row => row[0] == id);
    if (rowIndex === -1) {
      return res.status(404).json({ error: 'ID no encontrado', status: false });
    }

    const updatedRange = `ANEXOS_TEC!A${rowIndex + 1}:M${rowIndex + 1}`;
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

//Función para eliminar los anexos de la hoja de calculo de Docencia y Servicio
router.post('/deleteAnexo', async (req, res) => {
  try {
      const { id } = req.body;
      
      if (!id) {
          return res.status(400).json({ error: 'ID faltante' });  // Verifica si el ID está presente en el request
      }
      
      const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
      const range = 'ANEXOS_TEC!A1:M1000';
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

      const emptyRow = ['', '', '', '', '', '', '', '', '', '', '', '', ''];
      const updateRange = `ANEXOS_TEC!A${rowIndex + 1}:M${rowIndex + 1}`;

      await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: updateRange,
          valueInputOption: 'RAW',
          resource: { values: [emptyRow] },
          key: process.env.key,
      });

      res.status(200).json({ success: 'Anexo eliminado correctamente', status: true });
  } catch (error) {
      console.error('Error en la eliminación:', error);
      res.status(400).json({ error: 'Error en la eliminación', status: false });
  }
});

//Función para enviar los datos de la práctica a la hoja de cálculo de Docencia y Servicio
router.post('/sendPractice', async (req, res) => {
  try {
    const { insertData, sheetName } = req.body;
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    const range = sheetName;
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    
    // Obtener los datos actuales para calcular la siguiente fila disponible
    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    const currentValues = responseSheet.data.values;
    const nextRow = currentValues ? currentValues.length + 1 : 1; // Calcular la siguiente fila
    const updatedRange = `${range}!A${nextRow}`;
    
    // Insertar los datos en la siguiente fila
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: { values: insertData },
      key: process.env.key,
    });

    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Solicitud guardada correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'Error al guardar la solicitud', status: false });
    }
  } catch (error) {
    console.error('Error en la conexión:', error);
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

//Función para actualizar los datos de la práctica de la hoja de cálculo de Docencia y Servicio
router.post('/updatePractice', async (req, res) => {
  try {
    const { updateData, id, sheetName } = req.body;
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    const range = `${sheetName}!A1:I1000`; // Asegúrate de tener el rango correcto
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    const currentValues = responseSheet.data.values;

    // Encontrar la fila correspondiente al ID
    const rowIndex = currentValues.findIndex(row => row[0] == id);
    if (rowIndex === -1) {
      return res.status(404).json({ error: 'ID no encontrado', status: false });
    }

    // Actualizar la fila
    const updatedRange = `${sheetName}!A${rowIndex + 1}:I${rowIndex + 1}`;
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: { values: [updateData] },
      key: process.env.key,
    });

    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Solicitud actualizada correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'Error al actualizar la solicitud', status: false });
    }
  } catch (error) {
    console.error('Error en la conexión:', error);
    return res.status(400).json({ error: 'Error en la conexión', status: false });
  }
});

//Función para eliminar los datos de la práctica de la hoja de cálculo de Docencia y Servicio
router.post('/deletePractice', async (req, res) => {
  try {
    const { id, sheetName } = req.body;
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    const range = `${sheetName}!A1:I1000`; // Rango de la hoja donde buscar
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });

    console.log("Intentando eliminar el ID:", id);

    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });

    const currentValues = responseSheet.data.values;

    // Buscar el índice de la fila por el ID
    const rowIndex = currentValues.findIndex(row => row[0] == id);
    if (rowIndex === -1) {
      console.error('ID no encontrado:', id);
      return res.status(404).json({ error: 'ID no encontrado', status: false });
    }

    // Actualizar la fila con valores vacíos (efecto de eliminación)
    const updatedRange = `${sheetName}!A${rowIndex + 1}:I${rowIndex + 1}`;
    const updateResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: { values: [['']] }, // Rellenar la fila con valores vacíos
      key: process.env.key,
    });

    if (updateResponse.status === 200) {
      console.log('Fila eliminada (rellenada con valores vacíos) correctamente');
      res.status(200).json({ success: 'Solicitud eliminada correctamente', status: true });
    } else {
      console.error('Error al eliminar la solicitud');
      res.status(400).json({ error: 'Error al eliminar la solicitud', status: false });
    }
  } catch (error) {
    console.error('Error al eliminar la solicitud:', error);
    res.status(400).json({ error: 'Error al eliminar la solicitud', status: false });
  }
});

//Función para obtener los datos de la hoja de cálculo de Docencia y Servicio
router.post('/getPractices', async (req, res) => {
  try {
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y'; // ID de tu hoja
    const range = 'SOLICITUD_PRACT!A1:I1000';  // Rango donde están almacenadas las prácticas

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });

    res.json({
      status: true,
      data: sheetValuesToObject(response.data.values)  // Convierte los valores de la hoja en objetos
    });
  } catch (error) {
    console.log('Error en la solicitud:', error);
    res.json({
      status: false,
      error: 'Error en la solicitud'
    });
  }
});

//Función para obtener los escenarios de práctica de la hoja ESC_PRACTICA
router.post('/getInstituciones', async (req, res) => {
  try {
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    const range = 'ESC_PRACTICA!A1:F1000'; // Columnas A (id), B (nombre), C (tipología), D (código), E (fecha_inicio), F (fecha_fin)

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });

    // Debug: mostrar qué datos están llegando de la hoja
    console.log('=== DEBUG ESC_PRACTICA BACKEND ===');
    console.log('Headers raw (primera fila):', response.data.values[0]);
    console.log('Segunda fila ejemplo:', response.data.values[1]);
    console.log('Headers normalizados:', response.data.values[0]?.map(h => h.toLowerCase().trim()));
    
    const processedData = sheetValuesToObject(response.data.values);
    console.log('Primer objeto procesado:', processedData[0]);
    console.log('Campos disponibles:', Object.keys(processedData[0] || {}));
    console.log('==================================');

    res.json({
      status: true,
      data: processedData
    });
  } catch (error) {
    console.log('Error al obtener escenarios de práctica:', error);
    res.json({
      status: false,
      error: 'Error al obtener escenarios de práctica'
    });
  }
});

//Función para enviar documentos de escenario a la hoja ANEXOS_ESC
router.post('/sendDocEscenario', async (req, res) => {
  try {
    const { insertData, sheetName } = req.body;
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    const range = sheetName;
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    
    // Obtener los datos actuales para calcular la siguiente fila disponible
    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: process.env.key,
    });
    const currentValues = responseSheet.data.values;
    const nextRow = currentValues ? currentValues.length + 1 : 1;
    const updatedRange = `${range}!A${nextRow}`;
    
    // Insertar los datos en la siguiente fila
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW',
      resource: { values: insertData },
      key: process.env.key,
    });

    if (sheetsResponse.status === 200) {
      console.log('Documento de escenario guardado correctamente');
      return res.status(200).json({ success: 'Documento de escenario guardado correctamente', status: true });
    } else {
      return res.status(400).json({ error: 'Error al guardar el documento de escenario', status: false });
    }
  } catch (error) {
    console.error('Error al enviar documento de escenario:', error);
    return res.status(400).json({ error: 'Error al enviar documento de escenario', status: false });
  }
});

//Función para obtener los documentos de escenario de la hoja ANEXOS_ESC
router.post('/getDocEscenarios', async (req, res) => {
  try {
    const sheets = google.sheets({ version: 'v4', auth: jwtClient });
    const spreadsheetId = '1hPcfadtsMrTOQmH-fDqk4d1pPDxYPbZ712Xv4ppEg3Y';
    const range = 'ANEXOS_ESC!A1:I1000'; // Columnas A-I (id, id_programa, id_escenario, institucion, url, tipologia, codigo, fecha_inicio, fecha_fin)

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
    console.log('Error al obtener documentos de escenario:', error);
    res.json({
      status: false,
      error: 'Error al obtener documentos de escenario'
    });
  }
});

app.use(express.json());
app.use(bodyParser.json());
app.use(cors());
app.use(router);

// Servir archivos estáticos desde la carpeta 'uploads'
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

app.listen(PORT, () => {
  console.log(`Servidor backend escuchando en el puerto ${PORT}`);
});
