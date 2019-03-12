//Paquetes:
var azure = require('azure-storage');
var Excel = require('exceljs');

//Crear conexión:
var azure2 = require('./keys_azure'); //Importación de llaves
var tableSvc = azure.createTableService(azure2.myaccount, azure2.myaccesskey);

//Variables para realizar una busqueda:
var tablaUsar = "botdyesatb02"

//Query:
var query = new azure.TableQuery()
    .select(['PartitionKey', 'RowKey', 'Proyecto', 'Borrado', 'Check', 'Resguardo']);
var nextContinuationToken = null;
var nextContinuationToken2 = null;

//Contador:
var aceptCount = 0;
var count3 = 0;
var count2 = 0;
var count1 = 0;
var count0 = 0;
var total = 0;

//Celda nueva trabajando:
var celdaActual = 1;

//Crear Libro Final:
var workbookFinal = new Excel.Workbook('Name');
var worksheet = workbookFinal.addWorksheet('Hoja1');

//Trabajo promesa:
function promesaWork() {
    return new Promise(function(resolve, reject) { //Promesa 1

        //Colocación del titulo de cada columna:
        worksheet.getCell(`A${celdaActual}`).value = 'RowKey';
        worksheet.getCell(`B${celdaActual}`).value = 'PartitionKey';
        worksheet.getCell(`C${celdaActual}`).value = 'Proyecto';
        worksheet.getCell(`D${celdaActual}`).value = 'Borrado';
        worksheet.getCell(`E${celdaActual}`).value = 'Check';
        worksheet.getCell(`F${celdaActual}`).value = 'Resguardo';

        tableSvc.queryEntities(tablaUsar, query, null, function(error, results, response) {
            if (!error) {
                //Recorrido por row:
                //console.log(results);
                results.entries.forEach(function(entry) {
                    //Logica por row:
                    aceptCount = 0;
                    //console.log(`${entry['RowKey']['_']}`);
                    if (entry['Borrado']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['Check']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['Resguardo']['_'] == "Aprobado") {
                        aceptCount++;
                    }

                    console.log(`Aceptados: ${aceptCount}`);

                    if (aceptCount == 3) {
                        //Colocación del titulo de cada columna:
                        celdaActual++;
                        worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                        worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                        worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                        worksheet.getCell(`D${celdaActual}`).value = entry['Borrado']['_'];
                        worksheet.getCell(`E${celdaActual}`).value = entry['Check']['_'];
                        worksheet.getCell(`F${celdaActual}`).value = entry['Resguardo']['_'];
                        count3++;
                    } else if (aceptCount == 2) {
                        count2++;
                    } else if (aceptCount == 1) {
                        count1++;
                    } else {
                        count0++;
                    }
                });
            }

            //Token que permite continuar despues de leer 1000 rows:
            if (results.continuationToken) {
                nextContinuationToken = results.continuationToken;
            }

            resolve();
        });
    }).then(() => {
        return new Promise(function(resolve, reject) { // Promesa 2
            if (nextContinuationToken != null) {
                tableSvc.queryEntities(tablaUsar, query, nextContinuationToken, function(error, results, response) {
                    if (!error) {
                        //Recorrido por row:
                        //console.log(results);
                        results.entries.forEach(function(entry) {
                            //Logica por row:
                            aceptCount = 0;
                            //console.log(`${count}: ${entry['Borrado']['_']}`);
                            if (entry['Borrado']['_'] == "Aprobado") {
                                aceptCount++;
                            }
                            if (entry['Check']['_'] == "Aprobado") {
                                aceptCount++;
                            }
                            if (entry['Resguardo']['_'] == "Aprobado") {
                                aceptCount++;
                            }

                            console.log(`Aceptados: ${aceptCount}`);

                            if (aceptCount == 3) {
                                //Colocación del titulo de cada columna:
                                celdaActual++;
                                worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                                worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                                worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                                worksheet.getCell(`D${celdaActual}`).value = entry['Borrado']['_'];
                                worksheet.getCell(`E${celdaActual}`).value = entry['Check']['_'];
                                worksheet.getCell(`F${celdaActual}`).value = entry['Resguardo']['_'];
                                count3++;
                            } else if (aceptCount == 2) {
                                count2++;
                            } else if (aceptCount == 1) {
                                count1++;
                            } else {
                                count0++;
                            }
                        });
                        //Token que permite continuar despues de leer 1000 rows:
                        if (results.continuationToken) {
                            nextContinuationToken = results.continuationToken;
                        }
                        resolve();
                    }
                });
            }
        }).then(() => {
            return new Promise(function(resolve, reject) { //Promesa 3 //Copea desde aquí
                if (nextContinuationToken != null) {
                    tableSvc.queryEntities(tablaUsar, query, nextContinuationToken, function(error, results, response) {
                        if (!error) {
                            //Recorrido por row:
                            //console.log(results);
                            results.entries.forEach(function(entry) {
                                //Logica por row:
                                aceptCount = 0;
                                //console.log(`${count}: ${entry['Borrado']['_']}`);
                                if (entry['Borrado']['_'] == "Aprobado") {
                                    aceptCount++;
                                }
                                if (entry['Check']['_'] == "Aprobado") {
                                    aceptCount++;
                                }
                                if (entry['Resguardo']['_'] == "Aprobado") {
                                    aceptCount++;
                                }

                                console.log(`Aceptados: ${aceptCount}`);

                                if (aceptCount == 3) {
                                    //Colocación del titulo de cada columna:
                                    celdaActual++;
                                    worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                                    worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                                    worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                                    worksheet.getCell(`D${celdaActual}`).value = entry['Borrado']['_'];
                                    worksheet.getCell(`E${celdaActual}`).value = entry['Check']['_'];
                                    worksheet.getCell(`F${celdaActual}`).value = entry['Resguardo']['_'];
                                    count3++;
                                } else if (aceptCount == 2) {
                                    count2++;
                                } else if (aceptCount == 1) {
                                    count1++;
                                } else {
                                    count0++;
                                }
                            });
                            //Token que permite continuar despues de leer 1000 rows:
                            if (results.continuationToken) {
                                nextContinuationToken = results.continuationToken;
                            }
                            resolve();
                        }
                    });
                }
            }).then(() => {
                return new Promise(function(resolve, reject) { //Promesa 4 -----------------------------------------
                    if (nextContinuationToken != null) {
                        tableSvc.queryEntities(tablaUsar, query, nextContinuationToken, function(error, results, response) {
                            if (!error) {
                                //Recorrido por row:
                                //console.log(results);
                                results.entries.forEach(function(entry) {
                                    //Logica por row:
                                    aceptCount = 0;
                                    //console.log(`${count}: ${entry['Borrado']['_']}`);
                                    if (entry['Borrado']['_'] == "Aprobado") {
                                        aceptCount++;
                                    }
                                    if (entry['Check']['_'] == "Aprobado") {
                                        aceptCount++;
                                    }
                                    if (entry['Resguardo']['_'] == "Aprobado") {
                                        aceptCount++;
                                    }

                                    console.log(`Aceptados: ${aceptCount}`);

                                    if (aceptCount == 3) {
                                        //Colocación del titulo de cada columna:
                                        celdaActual++;
                                        worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                                        worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                                        worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                                        worksheet.getCell(`D${celdaActual}`).value = entry['Borrado']['_'];
                                        worksheet.getCell(`E${celdaActual}`).value = entry['Check']['_'];
                                        worksheet.getCell(`F${celdaActual}`).value = entry['Resguardo']['_'];
                                        count3++;
                                    } else if (aceptCount == 2) {
                                        count2++;
                                    } else if (aceptCount == 1) {
                                        count1++;
                                    } else {
                                        count0++;
                                    }
                                });
                                //Token que permite continuar despues de leer 1000 rows:
                                if (results.continuationToken) {
                                    nextContinuationToken = results.continuationToken;
                                }
                                resolve();

                                //"Pega aqui abajo el codigo de resultado"

                            }
                        });
                    }
                }).then(() => {
                    console.log("Esto se lee la final."); //Hasta aquí --------------------------------------------
                });
            });
        });
    });
}

//Codigo de resultado:
//Pega esto en la promesa final, justo abajo del  resolve();
console.log(`${count3} tienen los 3 campos Aprobados`);
console.log(`${count2} tienen los 2 campos Aprobados`);
console.log(`${count1} tienen 1 campo Aprobado`);
console.log(`${count0} no tienen ningun Aprobado`);
total = count0 + count1 + count2 + count3;
console.log(`Total de campos analizados: ${total}`);

if (celdaActual > 1) {
    workbookFinal.xlsx.writeFile('finalAprobadoTabla02.xlsx').then(function() { //Puedes colocar cualquier nombre al archivo final sustituyendo "final.xlsx" (recuerda respetar siempre la extención .xlsx).
        console.log("¡El archivo se a creado correctamente!");
    });
} else {
    console.log("No hay información para crear el archivo");
}
//---------------------------------------------------------

promesaWork();