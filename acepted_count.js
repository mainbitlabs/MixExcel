//Rutas:
var libroPath = "./libroOrigenBot.csv"; //Libro a revisar:

//Paquetes:
//Con npm, instala los siguentes paquetes:
//npm install xlsx
//npm install exceljs
const XLSX = require('xlsx');
var Excel = require('exceljs');

//Leer libro de origen:
var libroOrigen = XLSX.readFile(libroPath);
var sheet_name_list = libroOrigen.SheetNames;
dataBook = XLSX.utils.sheet_to_json(libroOrigen.Sheets[sheet_name_list[0]]);

//Crear el nuevo libro:
var workbookFinal = new Excel.Workbook();
var worksheet = workbookFinal.addWorksheet();

//Control de celda nueva:
var celdaActual = 1;
var aceptCount = 0;

//Control del conteo "Aprobado":
var isAcepted = false;
var acptedDocument = 0;

//Programa:
function working(aprobadosRequeridos) {

    //Colocación del titulo de cada columna:
    worksheet.getCell(`A${celdaActual}`).value = "RowKey";
    worksheet.getCell(`B${celdaActual}`).value = "PartitionKey";
    worksheet.getCell(`C${celdaActual}`).value = "Proyecto";
    worksheet.getCell(`D${celdaActual}`).value = "Borrado";
    worksheet.getCell(`E${celdaActual}`).value = "Check";
    worksheet.getCell(`F${celdaActual}`).value = "Resguardo";
    celdaActual++;

    //Analizando libro de origen:
    for (var key in dataBook) {
        //Reseteo del conteo "Aprobado"
        aceptCount = 0;
        isAcepted = false;

        //Conteo de aprobados en una fila:
        if (dataBook[key]['Borrado'] == "Aprobado") {
            aceptCount++;
        }
        if (dataBook[key]['Check'] == "Aprobado") {
            aceptCount++;
        }
        if (dataBook[key]['Resguardo'] == "Aprobado") {
            aceptCount++;
        }

        //Dependiendo del numero de aprobados que se querieran, se creara o no la nueva fila en Excel:
        if (aceptCount == aprobadosRequeridos) {
            isAcepted = true;
            acptedDocument++;

            worksheet.getCell(`A${celdaActual}`).value = `${dataBook[key]['RowKey']}`;
            worksheet.getCell(`B${celdaActual}`).value = `${dataBook[key]['PartitionKey']}`
            worksheet.getCell(`C${celdaActual}`).value = `${dataBook[key]['Proyecto']}`
            worksheet.getCell(`D${celdaActual}`).value = `${dataBook[key]['Borrado']}`
            worksheet.getCell(`E${celdaActual}`).value = `${dataBook[key]['Check']}`
            worksheet.getCell(`F${celdaActual}`).value = `${dataBook[key]['Resguardo']}`

            celdaActual++;
        }
        console.log(`${dataBook[key]['RowKey']} con aprobado: ${aceptCount} - ${isAcepted}`);
    }
    console.log(`Hay ${acptedDocument} con ${aprobadosRequeridos} aprobados.`);

    //Guardando nuevo libro:
    workbookFinal.xlsx.writeFile('finalbot2Apro.xlsx').then(function() {
        console.log("saved");
    });
}

//Ejecuta el programa y coloca como parametro el numero de aprobas que requieres en una fila para separar esta del resto:
working(2);