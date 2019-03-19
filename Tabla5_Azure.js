//Rutas:
var libroAzulPath = "./LibroAzul.xlsx"; //Ruta libro azul

//Paquetes:
var azure = require('azure-storage');
const XLSX = require('xlsx');
var Excel = require('exceljs');

//Crear conexión:
var azure2 = require('./keys_azure'); //Importación de llaves
var tableSvc = azure.createTableService(azure2.myaccount, azure2.myaccesskey);

//Tabla origen:
var tablaUsar = "botdyesatb05"

//Leer Libro Azul:
var workbookAzul = XLSX.readFile(libroAzulPath);
var sheet_name_list = workbookAzul.SheetNames;
dataAzul = XLSX.utils.sheet_to_json(workbookAzul.Sheets[sheet_name_list[0]]);

//Crear Libro 3Aprobados:
var workbook3Aprobados = new Excel.Workbook('algo');
var worksheet3Aprobados = workbook3Aprobados.addWorksheet('Hoja1');

//Crear Libro NoCumple:
var workbookNoCumple = new Excel.Workbook('algomas');
var worksheetNoCumple = workbookNoCumple.addWorksheet('Hoja2');

//Query:
var query = new azure.TableQuery();
var nextContinuationToken = null;

//Variables:
var taskTabla5 = {
    PartitionKey: { '_': '' },
    RowKey: { '_': '' },
    Timestamp: { '_': '' },
    Baja: { '_': '' },
    Borrado: { '_': '' },
    Check: { '_': '' },
    Resguardo: { '_': '' },
    HojaDeServicio: { '_': '' },
    Status: { '_': '' }
};
var contador = 0;
var contadorRepeticion = 0;
var contadorAprobados = 0;
var borrado = "";
var check = "";
var resguardo = "";
var celdaActual3Aprobados = 1;
var celdaActualNoCumple = 1;
var iguales = 0;
var finalizar = false;

//Programa
async function working() {
    //Colocación del titulo de cada columna del libro 3Aprobados:
    worksheet3Aprobados.getCell(`A${celdaActual3Aprobados}`).value = 'RowKey';
    worksheet3Aprobados.getCell(`B${celdaActual3Aprobados}`).value = 'Borrado';
    worksheet3Aprobados.getCell(`C${celdaActual3Aprobados}`).value = 'Check';
    worksheet3Aprobados.getCell(`D${celdaActual3Aprobados}`).value = 'Resguardo';

    //Colocación del titulo de cada columna del libro NoCumple:
    worksheetNoCumple.getCell(`A${celdaActualNoCumple}`).value = 'RowKey';
    worksheetNoCumple.getCell(`B${celdaActualNoCumple}`).value = 'Borrado';
    worksheetNoCumple.getCell(`C${celdaActualNoCumple}`).value = 'Check';
    worksheetNoCumple.getCell(`D${celdaActualNoCumple}`).value = 'Resguardo';

    celdaActual3Aprobados++;
    celdaActualNoCumple++;

    //Bucle:
    do {
        await promesa();
    } while (finalizar == false);
    resultado();
}

function promesa() {
    return new Promise(function(resolve, reject) { //Promesa 1
        tableSvc.queryEntities(tablaUsar, query, nextContinuationToken, function(error, results, response) {

            //Logica por cada entidad:
            if (!error) {
                results.entries.forEach(function(entry) {
                    contador++;
                    borrado = "";
                    check = "";
                    resguardo = "";
                    contadorRepeticion = 0;
                    contadorAprobados = 0;
                    console.log(`${entry['RowKey']['_']}`);
                    for (var key in dataAzul) { //Inicia comparación desde libro azul...
                        if (entry['RowKey']['_'] == dataAzul[key]['Serie']) {
                            if (dataAzul[key]['TipoDoc'] == "Borrado") {
                                borrado = dataAzul[key]['Estatus'];
                                if (dataAzul[key]['Estatus'] == "Aprobado") {
                                    contadorAprobados++;
                                }
                            } else if (dataAzul[key]['TipoDoc'] == "Check") {
                                check = dataAzul[key]['Estatus'];
                                if (dataAzul[key]['Estatus'] == "Aprobado") {
                                    contadorAprobados++;
                                }
                            } else if (dataAzul[key]['TipoDoc'] == "Resguardo") {
                                resguardo = dataAzul[key]['Estatus'];
                                if (dataAzul[key]['Estatus'] == "Aprobado") {
                                    contadorAprobados++;
                                }
                            }
                            console.log(`Son iguales: ${entry['RowKey']['_']} - ${dataAzul[key]['Serie']}`);
                            contadorRepeticion++;
                        }
                    }

                    if (contadorAprobados == 3) {
                        //Escribir en el libro 3Aprobados:
                        worksheet3Aprobados.getCell(`A${celdaActual3Aprobados}`).value = entry['RowKey']['_'];
                        worksheet3Aprobados.getCell(`B${celdaActual3Aprobados}`).value = borrado;
                        worksheet3Aprobados.getCell(`C${celdaActual3Aprobados}`).value = check;
                        worksheet3Aprobados.getCell(`D${celdaActual3Aprobados}`).value = resguardo;
                        celdaActual3Aprobados++;
                        iguales++;
                    } else {
                        //Colocación del titulo de cada columna del libro NoCumple:
                        worksheetNoCumple.getCell(`A${celdaActualNoCumple}`).value = entry['RowKey']['_'];
                        worksheetNoCumple.getCell(`B${celdaActualNoCumple}`).value = borrado;
                        worksheetNoCumple.getCell(`C${celdaActualNoCumple}`).value = check;
                        worksheetNoCumple.getCell(`D${celdaActualNoCumple}`).value = resguardo;
                        celdaActualNoCumple++;
                        iguales++;
                    }
                });
            }

            //Token que permite continuar despues de leer 1000 rows:
            if (results.continuationToken) {
                nextContinuationToken = results.continuationToken;
                resolve();
            } else {
                finalizar = true;
                resolve();
            }
        });
    });
}

//Funcion que se ejecuta el final del programa:
function resultado() {

    //Guardar libros:
    if (celdaActual3Aprobados > 1) {
        workbook3Aprobados.xlsx.writeFile('3Aprobados.xlsx').then(function() { //Puedes colocar cualquier nombre al archivo final sustituyendo "final.xlsx" (recuerda respetar siempre la extención .xlsx).
            console.log("¡El archivo 3Aprobados se a creado correctamente!");
        });
    } else {
        console.log("No hay información para crear el archivo");
    }

    if (celdaActualNoCumple > 1) {
        workbookNoCumple.xlsx.writeFile('NoCumple.xlsx').then(function() { //Puedes colocar cualquier nombre al archivo final sustituyendo "final.xlsx" (recuerda respetar siempre la extención .xlsx).
            console.log("¡El archivo NoCumple se a creado correctamente!");
        });
    } else {
        console.log("No hay información para crear el archivo");
    }

    console.log(`Se encontrarion ${contador} entidades.`);
    console.log(`Se encontrarion ${iguales} que coinciden.`);
    console.log("El programa ha terminado:");
}

//Inicia el trabajo:
working();