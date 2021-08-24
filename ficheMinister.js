const electron = require('electron');
const { ipcMain } = electron;
const XLSX = require('xlsx');
const moment = require('moment');
const mkdirp = require('mkdirp');
const path = require('path');
const abonne = XLSX.readFile('./ABONNE.DBF');
const abonment = XLSX.readFile('./ABONMENT.DBF');
const fact = XLSX.readFile('./FACTURES.DBF');
// const fact = XLSX.read("./FACTURES.DBF", { type: "buffer" });
// const fact = XLSX.utils.sheet_to_json("./FACTURES.DBF");
const minist = XLSX.readFile('./MINISTERS.DBF');
const commune = XLSX.readFile('./COMMUNE.DBF');
const unite = XLSX.readFile('./UNITE.DBF');
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
  'ADM',
  'MINIS',
];
function MinisterPerseXLSX() {
  return new Promise((resolve, reject) => {
    setTimeout(() => {
      const abonnerSheet = abonne.Sheets[abonne.SheetNames[0]];
      // console.log("ws:", ws);
      const ws_fact = fact.Sheets[fact.SheetNames[0]];
      const ws_minist = minist.Sheets[minist.SheetNames[0]];
      const ws_abonment = abonment.Sheets[abonment.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(abonnerSheet);
      const data_minist = XLSX.utils.sheet_to_json(ws_minist);
      const data_fact = XLSX.utils.sheet_to_json(ws_fact);
      const data_abomment = XLSX.utils.sheet_to_json(ws_abonment);
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
      // console.table(ws_facts);
      const ws_minists = data_minist.map((record) => {
        return [`${record.NUMAB}`, `${record.ADM}`, `${record.MINIS}`];
      });
      const ws_abnm = data_abomment.map((record) => {
        return [`${record.NUMAB}`, `${record.ETATCPT}`];
      });
      const ws_abn = data.map((record) => {
        return [`${record.NUMAB}`, `${record.RAISON}`, `${record.TYPABON}`];
      });
      const filaN = ws_abn.map((e) => {
        if (
          e[0] == ws_abnm[0] &&
          e[0] == ws_minists[0] &&
          e[0] == ws_facts[0]
        ) {
          return [
            `${e[0]}`,
            `${e[1]}`,
            `${e[2]}`,
            `${ws_abnm[1]}`,
            `${ws_minists[1]}`,
            `${ws_minists[2]}`,
            `${ws_facts[1]}`,
            `${ws_facts[2]}`,
            `${ws_facts[3]}`,
            `${ws_facts[4]}`,
            `${ws_facts[5]}`,
            `${ws_facts[6]}`,
            `${ws_facts[7]}`,
            `${ws_facts[8]}`,
          ];
        }
      });

      // const ws_NUMAB = data.map((rec) => {
      //   const ws_min = ws_minists.map((item) => {
      //     if (rec.CODRUE === item[0]) {
      //       console.log('yadra');
      //       const test = ws_abnm.map((element) => {
      //         if (rec.NUMAB === element[0]) {
      //           return [
      //             `${rec.NUMAB}`,
      //             `${rec.RAISON}`,
      //             `${rec.TYPABON}`,
      //             `${element[1]}`,
      //           ];
      //         }
      //         console.log('this');
      //       });
      //       const ff = ws_facts.map((wsFactu) => {
      //         if (rec.NUMAB === wsFactu[0]) {
      //           // console.log(`${rec.NUMAB}`);
      //           return [
      //             `${wsFactu[1]}`,
      //             `${wsFactu[2]}`,
      //             `${wsFactu[3]}`,
      //             `${wsFactu[4]}`,
      //             `${wsFactu[5]}`,
      //             `${wsFactu[6]}`,
      //             `${wsFactu[7]}`,
      //             `${wsFactu[8]}`,
      //             `${item[1]}`,
      //             `${item[2]}`,
      //           ];
      //         }
      //         console.log('here');
      //       });
      //       console.log('lehoooo');
      //       return Array.prototype.push.apply(test, ff);
      //     }
      //   });
      //   return ws_min;
      // });
      const newXlsx = XLSX.utils.book_new();
      const ws_name = `${unite_code} ${cell_unite}`;
      var wss = XLSX.utils.aoa_to_sheet(arr1.concat([ws_row], filaN));
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
    }, 1000);
  });
  // return console.log("hello");
}

module.exports.MinisterPerseXLSX = MinisterPerseXLSX;
