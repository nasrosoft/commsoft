// npm install --save-dev electron-rebuild
// hh .\node_modules\.bin\electron-rebuild.cmd

const electron = require('electron');

const { AbonnePerseXLSX } = require('./ficheAbonne');
const { MinisterPerseXLSX } = require('./ficheMinister');
const mkdirp = require('mkdirp');
const path = require('path');

const { app, BrowserWindow, ipcMain, Menu, dialog } = electron;

// const unite = XLSX.readFile("../UNITE.DBF");
let mainWindow;

const menuTemplate = [];
const createMenuwin = () => {
  const menu = Menu.buildFromTemplate(menuTemplate);
  Menu.setApplicationMenu(menu);
};

app.on('ready', () => {
  mainWindow = new BrowserWindow({
    height: 444, // 730 dev , 710 pro
    width: 320,
    title: 'UNITE DE BECHAR',
    resizable: false,
    // frame: false,
    webPreferences: {
      nodeIntegration: true,
      backgroundThrottling: false,
    },
  });

  mainWindow.loadURL(`file://${__dirname}/src/index.html`);
  createMenuwin();
  mainWindow.on('closed', () => app.quit());

  (async () => {
    // await copyData();
    await AbonnePerseXLSX();
    MinisterPerseXLSX().then(() => {
      console.log('azuul');
    });
  })();
});

if (process.env.NODE_ENV !== 'production') {
  menuTemplate.push({
    label: '++_DEVELOPER_++',
    submenu: [
      {
        label: 'Toggle Developer Tools',
        click(item, focusedWindow) {
          focusedWindow.toggleDevTools();
        },
      },
    ],
  });
}

// fs.copyFile('ABONNE.DBF', 'ABONNE.xls', (err) => {
//   if (err) throw err;
//   // const abonne = XLSX.readFile('ABONNE.xls');
//   // const ws = abonne.Sheets["ABONNE"];
//   // const data = XLSX.utils.sheet_to_json(ws);
//   console.log('ABONNE File was copied to destination');

// });
// fs.copyFile('RUE.DBF', 'RUE.xls', (err) => {
//   if (err) throw err;
//   // const rue = XLSX.readFile('RUE.xls');
//   // const ws_rue = rue.Sheets["RUE"];
//   // const data_rue = XLSX.utils.sheet_to_json(ws_rue);
//   console.log('RUE File was copied to destination');
// });
// module.exports.renderAbonner = renderAbonner;
