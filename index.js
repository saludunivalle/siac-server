const express = require('express');
const { google } = require('googleapis');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const stream = require("stream"); // Added
const multer = require('multer');
const fs = require('fs');
const {sheetValuesToObject} = require('./utils');
const { config } = require('dotenv');
config();

const app = express();
const router = express.Router();
const PORT = process.env.PORT || 3001;
console.log(process.env.client_email, process.env.private_key)
//configure a JWT auth client
let jwtClient = new google.auth.JWT(
  process.env.client_email,
  null,
  (process.env.private_key).replace(/\\n/g, '\n'),
  ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']);

  //authenticate request
  jwtClient.authorize(function (err, tokens) {
        if (err) {
        console.log(err,'hear');
        return;
  } else {
       console.log("Successfully connected!");
  }
});

app.use(bodyParser.json());

router.post('/sendData', async ( req, res) => {
  
  try {
    const {
      insertData, 
      sheetName
    }=req.body
    const spreadsheetId = '1GQY2sWovU3pIBk3NyswSIl_bkCi86xIwCjbMqK_wIAE';
    const range = sheetName;
    const sheets = google.sheets({ version: 'v4' , auth: jwtClient });
    const responseSheet = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      key: 'AIzaSyDQTWi9NHU_UTjhVQ1Wb08qxREaRgD9v1g',
    });
    const currentValues = responseSheet.data.values;
    const nextRow = currentValues ? currentValues.length + 1 : 1;
    const updatedRange = `${range}!A${nextRow}`;
    const sheetsResponse = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: updatedRange,
      valueInputOption: 'RAW', 
      resource: {
        values: insertData
      },
      key: 'AIzaSyDQTWi9NHU_UTjhVQ1Wb08qxREaRgD9v1g',
    })
    if (sheetsResponse.status === 200) {
      return res.status(200).json({ success: 'Se inserto correctamente', status:true});
    } else {
      return res.status(400).json({ error: 'No se inserto', status:false});
    }
  } catch (error) {
    return res.status(400).json({ error: 'Error en la conexion', status:false});
  }
});

router.post('/', async ( req, res) => {
  
  try {
    // console.log(jwtClient);
    const sheets = google.sheets({ version: 'v4',  auth: jwtClient });
      const spreadsheetId = '1GQY2sWovU3pIBk3NyswSIl_bkCi86xIwCjbMqK_wIAE';
      //const range = 'PROGRAMAS';
      let range;
      switch (req.body.sheetName) {
        case 'Programas':
          range = 'PROGRAMAS!A1:AG93';
          break;
        case 'Seguimientos':
          range = 'SEGUIMIENTOS!A1:H100';
          break;
        case 'Permisos':
          range = 'PERMISOS!A1:C20';
          break;
        case 'Proc_Fases':
          range = 'PROC_FASES!A1:C100';
          break;
        case 'Proc_X_Prog':
          range = 'PROC_X_PROG!A1:C100';
          break;
        default:
          return res.status(400).json({ error: 'Nombre de hoja no válido' });
      }

      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
        key: 'AIzaSyDQTWi9NHU_UTjhVQ1Wb08qxREaRgD9v1g', 
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

//drive

const upload = multer();

router.post('/upload', upload.single('file'), async ( req, res) => {
  
  try {
    const drive = google.drive({
      version: 'v3',
      auth: jwtClient
    })

    const { path: filePath, originalname } = req.file;

    const fileMetadata = {
      name: originalname,
      parents: ['1Y5dit9wGoK9iKsOd7W-ACjee8zlMhI7R'] 
    };

    const templateBuffer = Buffer.from( req.file.buffer,'base64');

    const media = {
      mimeType: req.file.mimetype,
      body: new stream.PassThrough().end(templateBuffer),  
    };
  
    console.log('llego');
    console.log(req.file);
    const response = await drive.files.create({
        requestBody: fileMetadata,
        media: media
    });

    console.log('Archivo subido correctamente a Google Drive:', response.data);

    res.json({ success: true, message: 'Archivo subido correctamente a Google Drive', enlace: `https://drive.google.com/file/d/${response.data.id}/view`});

    // console.log(`Se recibió el archivo: ${originalname}`);
    // //console.log(drive);
    // console.log(req?.form?.files);
    // console.log(req?.files);
    // console.log(req.body);
    // res.json({files:''})
  } catch (error) {
    console.log(error);
    return res.status(400).json({ error:error, status:false});
  }
});

app.use(express.json());
app.use(bodyParser.json());
app.use(cors());
app.use(router)

app.listen(PORT, () => {
  console.log(`Servidor backend escuchando en el puerto ${PORT}`);
});
