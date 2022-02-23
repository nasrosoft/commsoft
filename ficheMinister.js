const electron = require('electron');
const { ipcMain } = electron;
const XLSX = require('xlsx');
const moment = require('moment');
const mkdirp = require('mkdirp');
const path = require('path');
const abonne = XLSX.readFile('./ABONNE.DBF');
const abonment = XLSX.readFile('./ABONMENT.DBF');
// const fact = XLSX.readFile('./FACTURES.DBF');
// const fact = XLSX.read("./FACTURES.DBF", { type: "buffer" });
// const fact = XLSX.utils.sheet_to_json("./FACTURES.DBF");
const minist = XLSX.readFile('./MINISTERS.DBF');
const commune = XLSX.readFile('./COMMUNE.DBF');
const unite = XLSX.readFile('./UNITE.DBF');
// ------------
const { DBFFile } = require('dbffile');

// ------------
// console.log("fact ::", fact);
// console.log("abonne ::", abonne);
const first_sheet_name_commune = commune.SheetNames[0];
const first_sheet_name_unite = unite.SheetNames[0];
// const address_of_cell = "D2";

const worksheet = commune.Sheets[first_sheet_name_commune];
// let commune_value = worksheet[address_of_cell].v;

const now = moment().utcOffset('+01:00');
const nowFileName = now.format('DMY_H-mm');

const cell_unite = unite.Sheets[first_sheet_name_unite]['B2'].v;
const unite_code = unite.Sheets[first_sheet_name_unite]['A2'].v;

let arr1 = [];

const ws_row = [
  'NUMAB',
  'ADM',
  'MINIS',
  'RAISON',
  'TYPABON',
  'ETATCPT',
  'TYPE',
  'DATFACT',
  'MONTTC',
  'PAIEMENT',
  'MODALITE',
  'DATREG',
  'DATSAISIE',
  'CHEQUE',
];
console.log('rani mena ');
function MinisterPerseXLSX() {
  console.log('rani mena -1');
  return new Promise((resolve, reject) => {
    console.log('rani mena 0');
    setTimeout(() => {
      const abonnerSheet = abonne.Sheets[abonne.SheetNames[0]];
      console.log('rani mena 1');
      // const ws_fact = fact.Sheets[fact.SheetNames[0]];
      console.log('rani mena 2');
      const ws_minist = minist.Sheets[minist.SheetNames[0]];
      const ws_abonment = abonment.Sheets[abonment.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(abonnerSheet);
      const data_minist = XLSX.utils.sheet_to_json(ws_minist);
      // const data_fact = XLSX.utils.sheet_to_json(ws_fact);
      const data_abomment = XLSX.utils.sheet_to_json(ws_abonment);
      // ================================================
      // *********** start
      async function testRead() {
        let dbf = await DBFFile.open('./FACTURES/FACTURES.DBF');
        console.log(`DBF file contains ${dbf.recordCount} records.`);
        console.log('dbf._recordLength', dbf._recordLength);
        // console.log(`Field names: ${dbf.fields.map((f) => f.name).join(', ')}`);
        let data_fact = await dbf.readRecords();
        // const data_fact = XLSX.utils.sheet_to_json(records);
        // console.log('records::', data_fact);
        const ws_facts = data_fact.map((record) => {
          return [
            `${record.NUMAB}`,
            `${record.TYPE}`,
            `${record.DATFACT}`,
            `${record.MONTTC}`,
            `${record.PAIEMENT}`,
            `${record.MODALITE}`,
            `${record.DATREG}`,
            `${record.DATSAISIE}`,
            `${record.CHEQUE}`,
          ];
        });

        // ***********  end

        const ws_minists = data_minist.map((record) => {
          return [`${record.NUMAB}`, `${record.ADM}`, `${record.MINIS}`];
        });
        const ws_abnm = data_abomment.map((record) => {
          return [`${record.NUMAB}`, `${record.ETATCPT}`];
        });
        const ws_abn = data.map((record) => {
          return [`${record.NUMAB}`, `${record.RAISOC}`, `${record.TYPABON}`];
        });
        console.log('abonne:', data.length);

        // =================================================

        const filaN = ws_minists
          .map((e) => {
            let result = [...e];
            for (let i = 0; i < ws_abn.length; i++) {
              if (e[0] === ws_abn[i][0] && Boolean(e[0])) {
                result = [...result, ws_abn[i][1], ws_abn[i][2]];
                break;
              }
            }

            for (let i = 0; i < ws_abnm.length; i++) {
              // console.log(`e[0] => ${e[0]} === ws_abnm[i][0] => ${ws_abnm[i][0]}`);
              if (e[0] === ws_abnm[i][0] && Boolean(e[0])) {
                result = [...result, ws_abnm[i][1]];
                break;
              }
            }

            return result;
          })
          .filter((e) => e.length > 0);

        const lastFinal = ws_facts
          .map((f) => {
            let res = [];
            for (let i = 0; i < filaN.length; i++) {
              if (f[0] === filaN[i][0] && Boolean(f[0])) {
                res = [
                  ...filaN[i],
                  f[1],
                  f[2],
                  f[3],
                  f[4],
                  f[5],
                  f[6],
                  f[7],
                  f[8],
                ];
              }
            }
            return res;
          })
          .filter((f) => f.length > 0);

        // console.log('filaN', lastFinal);
        // =================================================

        const newXlsx = XLSX.utils.book_new();
        const ws_name = `${unite_code} ${cell_unite}`;
        var wss = XLSX.utils.aoa_to_sheet(arr1.concat([ws_row], lastFinal));
        /* Add the worksheet to the workbook */
        XLSX.utils.book_append_sheet(newXlsx, wss, ws_name);
        ipcMain.on('hello', () => {
          let reqPath = path.join(__dirname, '../../');
          const made = mkdirp.sync(`${reqPath}/Fichier_Minister`);
          XLSX.writeFile(
            newXlsx,
            `${reqPath}/Fichier_Minister/${cell_unite}_${nowFileName}.xlsx`
          );
        });
        resolve();
      }
      testRead();
    }, 1000);
  });
}

module.exports.MinisterPerseXLSX = MinisterPerseXLSX;
