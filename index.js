const express = require('express');
const { google } = require('googleapis');
const bodyParser = require('body-parser');
const cors = require('cors');
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
  ['https://www.googleapis.com/auth/spreadsheets']);

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
          range = 'PROGRAMAS!A1:AE97';
          break;
        case 'Seguimientos':
          range = 'SEGUIMIENTOS!A1:G97';
          break;
        case 'Permisos':
          range = 'PERMISOS!A1:C20';
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

app.use(express.json());
app.use(bodyParser.json());
app.use(cors());
app.use(router)

app.listen(PORT, () => {
  console.log(`Servidor backend escuchando en el puerto ${PORT}`);
});
