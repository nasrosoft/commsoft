const electron = require("electron");
const { ipcMain } = electron;
const XLSX = require("xlsx");
const moment = require("moment");
const mkdirp = require("mkdirp");
const path = require("path");
const abonne = XLSX.readFile("./ABONNE.DBF");
const abonment = XLSX.readFile("./ABONMENT.DBF");
const rue = XLSX.readFile("./RUE.DBF");
const unite = XLSX.readFile("./UNITE.DBF");
const commune = XLSX.readFile("./COMMUNE.DBF");

const first_sheet_name_commune = commune.SheetNames[0];
const first_sheet_name_unite = unite.SheetNames[0];
const address_of_cell = "D2";

const worksheet = commune.Sheets[first_sheet_name_commune];
let commune_value = worksheet[address_of_cell].v;

const now = moment().utcOffset("+01:00");
const nowFileName = now.format("DMY_H-mm");
const cell_unite = unite.Sheets[first_sheet_name_unite]["B2"].v;
const unite_code = unite.Sheets[first_sheet_name_unite]["A2"].v;

let arr1 = [];

const ws_row = [
  "LIBCOM",
  "TYPABON",
  "NUMAB",
  "RAISOC",
  "CODRUE",
  "NOUVNOM",
  "DATECRE",
  "NUMCOMPT",
];
function AbonnePerseXLSX() {
  return new Promise((resolve, reject) => {
    setTimeout(() => {
      const ws = abonne.Sheets[abonne.SheetNames[0]];
      // console.log('ws:',ws);
      const ws_rue = rue.Sheets[rue.SheetNames[0]];
      const ws_abonment = abonment.Sheets[rue.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws);
      const data_rue = XLSX.utils.sheet_to_json(ws_rue);
      const data_abomment = XLSX.utils.sheet_to_json(ws_abonment);
      const ws_rues = data_rue.map((record) => {
        return [`${record.CODRUE}`, `${record.NOUVNOM}`];
      });
      const ws_abnm = data_abomment.map((record) => {
        return [`${record.NUMAB}`, `${record.NUMSER}`];
      });
      const ws_NUMAB = data.map((rec) => {
        for (const item of ws_rues) {
          if (rec.CODRUE === item[0]) {
            for (const element of ws_abnm) {
              if (rec.NUMAB === element[0]) {
                return [
                  `${commune_value}`,
                  `${rec.TYPABON}`,
                  `${rec.NUMAB}`,
                  `${rec.RAISOC}`,
                  `${rec.CODRUE}`,
                  `${item[1]}`,
                  `${rec.DATECRE}`,
                  `${element[1]}`,
                ];
              }
            }
          }
        }
      });
      const newXlsx = XLSX.utils.book_new();
      const ws_name = `${unite_code} ${cell_unite}`;
      var wss = XLSX.utils.aoa_to_sheet(arr1.concat([ws_row], ws_NUMAB));
      /* Add the worksheet to the workbook */
      XLSX.utils.book_append_sheet(newXlsx, wss, ws_name);
      ipcMain.on("hello", () => {
        let reqPath = path.join(__dirname, "../../");
        const made = mkdirp.sync(`${reqPath}/Fichier_Abonner`);
        XLSX.writeFile(
          newXlsx,
          `${reqPath}/Fichier_Abonner/${cell_unite}_${nowFileName}.xlsx`
        );
      });
      resolve();
    }, 3000);
  });
}

module.exports.AbonnePerseXLSX = AbonnePerseXLSX;
