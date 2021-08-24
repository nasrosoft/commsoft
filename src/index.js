// // https://www.npmjs.com/package/xlsx docs
const electron = require("electron");
const { ipcRenderer } = electron;

const btn = document.querySelector("#btn");
const btnMinister = document.querySelector("#btnMinister");

btn.addEventListener("click", () => {
  ipcRenderer.send("hello");
});

btnMinister.addEventListener("click", () => {
  ipcRenderer.send("ministaire");
});
