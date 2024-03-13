const express = require('express');
const { google } = require('googleapis');
const bodyParser = require('body-parser');
const cors = require('cors');
const {sheetValuesToObject, jwtClient} = require('./utils');

const app = express();
const router = express.Router();
const PORT = process.env.PORT || 3001;

app.use(bodyParser.json());

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
