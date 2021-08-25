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
console.log('rani mena ');
function MinisterPerseXLSX() {
  console.log('rani mena -1');
  return new Promise((resolve, reject) => {
    console.log('rani mena 0');
    setTimeout(() => {
      const abonnerSheet = abonne.Sheets[abonne.SheetNames[0]];
      console.log('rani mena 1');
      const ws_fact = fact.Sheets[fact.SheetNames[0]];
      console.log('rani mena 2');
      const ws_minist = minist.Sheets[minist.SheetNames[0]];
      const ws_abonment = abonment.Sheets[abonment.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(abonnerSheet);
      const data_minist = XLSX.utils.sheet_to_json(ws_minist);
      const data_fact = XLSX.utils.sheet_to_json(ws_fact);
      const data_abomment = XLSX.utils.sheet_to_json(ws_abonment);
      // ================================================
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
        return [`${record.NUMAB}`, `${record.RAISOC}`, `${record.TYPABON}`];
      });

      // =================================================
      const filaN = ws_facts.map((e) => {
        const obj = ws_abnm.find(
          (o) => ((o[0] === e[0]) === ws_minists[0][0]) === ws_abn[0][0]
        );
        return [
          `${ws_abn[0][0]}`,
          `${ws_abn[1][1]}`,
          `${ws_abn[2][2]}`,
          `${ws_abnm[1][1]}`,
          `${e[1]}`,
          `${e[2]}`,
          `${e[3]}`,
          `${e[4]}`,
          `${e[5]}`,
          `${e[6]}`,
          `${e[7]}`,
          `${e[8]}`,
          `${ws_minists[1][1]}`,
          `${ws_minists[2][2]}`,
        ];

        // if (e[0] === (ws_abnm[0][0] && ws_minists[0][0] && ws_facts[0][0])) {
        //   return [
        //     `${e[0]}`,
        //     `${e[1]}`,
        //     `${e[1]}`,
        //     `${ws_abnm[1][1]}`,
        //     `${ws_facts[1][1]}`,
        //     `${ws_facts[2][2]}`,
        //     `${ws_facts[3][3]}`,
        //     `${ws_facts[4][4]}`,
        //     `${ws_facts[5][5]}`,
        //     `${ws_facts[6][6]}`,
        //     `${ws_facts[7][7]}`,
        //     `${ws_facts[8][8]}`,
        //     `${ws_minists[1][1]}`,
        //     `${ws_minists[2][2]}`,
        //   ];
        // }
      });

      console.log(filaN);
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
