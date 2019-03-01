//Rutas:
const checklistFolder = './checklist_subir_BOT'; //Colocar ruta de carpeta "checklist_subir_BOT"
const resguardosFolder = './Resguardos_subir_BOT'; //Colocar ruta de carpeta "Resguardos_subir_BOT"
const csvFilePath = './Check14FebreroAcom.csv'; //Colocar ruta de archivo CSV a leer
const finalCheckPath = './finalCheck'; //Carpeta para los arichivos separados
const finalResguardosPath = './finalResguardos'; //Carpeta para los arichivos separados

//Paquetes:
//Con npm, instala los siguentes paquetes:
//npm install xlsx
//npm install exceljs
const fs = require('fs');
const XLSX = require('xlsx');

//Variables para separar nombres:
var comp1 = "";
var comp2 = "";

//Leer Excel:
var workbook = XLSX.readFile(csvFilePath);
var sheet_name_list = workbook.SheetNames;
data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

//Copiar archivos
function checklistFolderChange(ruta) {
    fs.readdir(ruta, (err, files) => {
        files.forEach(file => {
            comp1 = file.slice(0, -4);
            for (var key in data) {
                comp2 = data[key]['RowKey'];
                if (comp1 === comp2) {
                    fs.copyFile(`${checklistFolder}/${file}`, `${finalCheckPath}/${file}`, function(err) {
                        if (err) return console.error(err)
                    });
                    console.log("Igual");
                }
            }
        });
    });
}

function resguardosFolderChange(ruta) {
    fs.readdir(ruta, (err, files) => {
        files.forEach(file => {
            comp1 = file.slice(7, -4);
            for (var key in data) {
                comp2 = data[key]['RowKey'];
                if (comp1 === comp2) {
                    fs.copyFile(`${resguardosFolder}/${file}`, `${finalResguardosPath}/${file}`, function(err) {
                        if (err) return console.error(err)
                    });
                    console.log("Igual");
                }
            }
        });
    });
}

//Renombrar archivos:
function finalchecklistFolderChange(ruta) {
    fs.readdir(ruta, (err, files) => {
        files.forEach(file => {
            comp1 = file.slice(0, -4);
            for (var key in data) {
                comp2 = data[key]['RowKey'];
                if (comp1 === comp2) {
                    fs.rename(`${finalCheckPath}/${file}`, `${finalCheckPath}/${data[key]['NuevoNombre']}`, function(err) {
                        if (err) console.log('ERROR: ' + err);
                    });
                    console.log("Igual");
                }
            }
        });
    });
}

function finalresguardosFolderChange(ruta) {
    fs.readdir(ruta, (err, files) => {
        files.forEach(file => {
            comp1 = file.slice(7, -4);
            for (var key in data) {
                comp2 = data[key]['RowKey'];
                if (comp1 === comp2) {
                    fs.rename(`${finalResguardosPath}/${file}`, `${finalResguardosPath}/${data[key]['Content.Proyecto']}_${data[key]['RowKey']}_Resguardo_${data[key]['PartitionKey']}.pdf`, function(err) {
                        if (err) console.log('ERROR: ' + err);
                    });
                    console.log("Igual");
                }
            }
        });
    });
}

//Ejecutar programa:
//Para hacer trabajar este programa, descomenta una tarea y ejecuta el codigo.
//Una vez finalizada la tarea, vuelve a comentarla, quita las barras de la siguiente y ejecuta nuevamente el codigo.

//Tarea 1 (Copiar):
//checklistFolderChange(checklistFolder);
//resguardosFolderChange(resguardosFolder);

//Tarea 2 (Renombrar):
//finalchecklistFolderChange(finalCheckPath);
//finalresguardosFolderChange(finalResguardosPath);