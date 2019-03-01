//Rutas:
var libroAzulPath = "./LibroAzul.csv"; //Ruta libro azul
var libroVerdePath = "./LibroVerde.csv"; //Ruta libro verde

//Paquetes:
//Con npm, instala los siguentes paquetes:
//npm install xlsx
//npm install exceljs
const XLSX = require('xlsx');
var Excel = require('exceljs');

//Leer Libro Azul:
var workbookAzul = XLSX.readFile(libroAzulPath);
var sheet_name_list = workbookAzul.SheetNames;
dataAzul = XLSX.utils.sheet_to_json(workbookAzul.Sheets[sheet_name_list[0]]);

//Leer Libro Verde:
var workbookVerde = XLSX.readFile(libroVerdePath);
var sheet_name_list = workbookVerde.SheetNames;
dataVerde = XLSX.utils.sheet_to_json(workbookVerde.Sheets[sheet_name_list[0]]);

//Crear Libro Final:
var workbookFinal = new Excel.Workbook('algo');
var worksheet = workbookFinal.addWorksheet('Discography');

//Celda nueva trabajando:
var celdaActual = 1;
var rep = 0;

//Control libro azul:
var estatusBorrador = "";
var estatusCheck = "";
var estatusResguardos = "";


//Tabajo
function working() {

    //Colocación del titulo de cada columna:
    worksheet.getCell(`A${celdaActual}`).value = 'RowKey';
    worksheet.getCell(`B${celdaActual}`).value = 'ParitionKey';
    worksheet.getCell(`C${celdaActual}`).value = 'Borrado_Verde';
    worksheet.getCell(`D${celdaActual}`).value = 'Borrado_Azul';
    worksheet.getCell(`E${celdaActual}`).value = 'Check_Verde';
    worksheet.getCell(`F${celdaActual}`).value = 'Check_Azul';
    worksheet.getCell(`G${celdaActual}`).value = 'Resguardo_Verde';
    worksheet.getCell(`H${celdaActual}`).value = 'Resgurado_Azul';

    celdaActual++;

    for (var key in dataVerde) { //Inicia comparación desde libro verde...
        for (var key2 in dataAzul) { //Bucle para comprarar celda verde con todas las celdas azules...
            if (`${dataVerde[key]['RowKey']}` == `${dataAzul[key2]['RowKey']}`) {

                //En caso de encontrar una concidencia repetida:
                if (rep != key) {
                    console.log("---------------------------------------------------------------");
                    celdaActual++;

                    //Control libro azul:
                    estatusBorrador = "";
                    estatusCheck = "";
                    estatusResguardos = "";



                    rep = key;
                }

                console.log(`Igual: ${dataVerde[key]['RowKey']} - ${dataAzul[key2]['RowKey']}`);

                //Creando celdas principales:
                worksheet.getCell(`A${celdaActual}`).value = `${dataVerde[key]['RowKey']}`;
                worksheet.getCell(`B${celdaActual}`).value = `${dataVerde[key]['PartitionKey']}`;

                //Creando celdas desde el libro verde:
                worksheet.getCell(`C${celdaActual}`).value = `${dataVerde[key]['Borrado']}`;
                worksheet.getCell(`E${celdaActual}`).value = `${dataVerde[key]['Check']}`;
                worksheet.getCell(`G${celdaActual}`).value = `${dataVerde[key]['Resguardo']}`;

                //Creando celdas desde el libro azul:
                if (`${dataAzul[key2]['TipDoc']}` == "Borrado") {
                    if (estatusBorrador == "Aprobado") {
                        worksheet.getCell(`D${celdaActual}`).value = "Aprobado";
                        console.log("Aprobado repetido.");
                    } else {
                        worksheet.getCell(`D${celdaActual}`).value = `${dataAzul[key2]['Estatus']}`;
                        estatusBorrador = `${dataAzul[key2]['Estatus']}`;
                    }
                } else if (`${dataAzul[key2]['TipDoc']}` == "Check") {
                    if (estatusCheck == "Aprobado") {
                        worksheet.getCell(`F${celdaActual}`).value = "Aprobado";
                        console.log("Aprobado repetido.");
                    } else {
                        worksheet.getCell(`F${celdaActual}`).value = `${dataAzul[key2]['Estatus']}`;
                        estatusCheck = `${dataAzul[key2]['Estatus']}`;
                    }
                } else if (`${dataAzul[key2]['TipDoc']}` == "Resguardo") {
                    if (estatusResguardos == "Aprobado") {
                        worksheet.getCell(`H${celdaActual}`).value = "Aprobado";
                        console.log("Aprobado repetido.");
                    } else {
                        worksheet.getCell(`H${celdaActual}`).value = `${dataAzul[key2]['Estatus']}`;
                        estatusResguardos = `${dataAzul[key2]['Estatus']}`;
                    }
                }

            }
        }
    }

    //Guardando nuevo libro:
    workbookFinal.xlsx.writeFile('final.xlsx').then(function() { //Puedes colocar cualquier nombre al archivo final sustituyendo "final.xlsx" (recuerda respetar siempre la extención .xlsx).
        console.log("¡El archivo se a creado correctamente!");
    });
}

//Ejecutando el programa:
working();