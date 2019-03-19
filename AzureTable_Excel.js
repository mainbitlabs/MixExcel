//Paquetes:
var azure = require('azure-storage');

//Crear conexión:
var azure2 = require('./keys_azure'); //Importación de llaves
var tableSvc = azure.createTableService(azure2.myaccount, azure2.myaccesskey);

//Tabla origen:
var tablaUsar = "botdyesatb02"
var proyectoTrabajando = "HOLIDAY-IMP";
var updateTaskTabla2 = {
    PartitionKey: { '_': '' },
    RowKey: { '_': '' },
    Timestamp: { '_': '' },
    Area: { '_': '' },
    Baja: { '_': '' },
    Borrado: { '_': '' },
    Check: { '_': '' },
    Descripcion: { '_': '' },
    Fecha_Fin: { '_': '' },
    Fecha_ini: { '_': '' },
    HojaDeServicio: { '_': '' },
    Inmueble: { '_': '' },
    Localidad: { '_': '' },
    NombreEnlace: { '_': '' },
    NombreUsuario: { '_': '' },
    Pospuesto: { '_': '' },
    Proyecto: { '_': '' },
    Resguardo: { '_': '' },
    SerieBorrada: { '_': '' },
    Servicio: { '_': '' },
    Status: { '_': '' },
};

//JSON tabla4:
var tablaUsar4 = "botdyesatb04"
var taskTabla4 = {
    PartitionKey: { '_': 'Proyecto' },
    RowKey: { '_': 'NombreProyecto' },
    Timestamp: { '_': '' },
    NumDoc: { '_': 0 },
    Baja: { '_': '' },
    Borrado: { '_': '' },
    Check: { '_': '' },
    Resguardo: { '_': '' },
    HojaDeServicio: { '_': '' },
};

//JSON tabla5:
var tablaUsar5 = "botdyesatb05"
var taskTabla5 = {
    PartitionKey: { '_': 'Proyecto' },
    RowKey: { '_': 'Serie' },
    Timestamp: { '_': '' },
    Baja: { '_': '' },
    Borrado: { '_': '' },
    Check: { '_': '' },
    Resguardo: { '_': '' },
    HojaDeServicio: { '_': '' },
    ColNueva: { '_': '' }
};

//Query:
var query = new azure.TableQuery();
var nextContinuationToken = null;
var numeroBucle = 1;
var tiempo = 5000;

//Contador:
var aceptCount = 0;
var proyectoCount = 0;
var count5 = 0;
var count4 = 0;
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
        tableSvc.queryEntities(tablaUsar, query, null, function(error, results, response) {
            if (!error) {
                //Recorrido por row:
                //console.log(results);
                results.entries.forEach(function(entry) {
                    //Logica por row:
                    aceptCount = 0;
                    //console.log(`${entry['RowKey']['_']}`);
                    if (entry['Baja']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['Borrado']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['Check']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['Resguardo']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['HojaDeServicio']['_'] == "Aprobado") {
                        aceptCount++;
                    }

                    console.log(`Aceptados: ${aceptCount}`);
                    if (aceptCount == 5) {
                        count5++;
                    } else if (aceptCount == 4) {
                        count4++;
                    } else if (aceptCount == 3) {
                        count3++;
                    } else if (aceptCount == 2) {
                        count2++;
                    } else if (aceptCount == 1) {
                        count1++;
                    } else {
                        //Colocación de la información:
                        if (entry['Proyecto']['_'] == proyectoTrabajando) {
                            proyectoCount++;
                            taskTabla5['PartitionKey']['_'] = entry['Proyecto']['_'];
                            taskTabla5['RowKey']['_'] = entry['RowKey']['_'];
                            taskTabla5['Baja']['_'] = entry['Baja']['_'];
                            taskTabla5['Borrado']['_'] = entry['Borrado']['_'];
                            taskTabla5['Check']['_'] = entry['Check']['_'];
                            taskTabla5['Resguardo']['_'] = entry['Resguardo']['_'];
                            taskTabla5['HolaDeServicio']['_'] = entry['HolaDeServicio']['_'];
                            tableSvc.insertEntity(tablaUsar5, taskTabla5, function(error, result, response) {
                                if (!error) {
                                    console.log("La entidad se agrego correctamente a la tabla 5");
                                }
                            });
                            count0++;
                        }
                    }
                });
            }

            //Token que permite continuar despues de leer 1000 rows:
            if (results.continuationToken) {
                nextContinuationToken = results.continuationToken;
                resolve();
            } else {
                return;
            }
        });
    }).then(() => {
        return new Promise(async function(resolve, reject) { //Promesa 2
            var continuar = bucleQuery();
            setTimeout(function() {
                if (continuar != false) {
                    resolve();
                } else {
                    resultado();
                }
                numeroBucle++;
                console.log(`Bucle numero: ${numeroBucle}`);
            }, tiempo);
        }).then(() => {
            return new Promise(async function(resolve, reject) { //Promesa 3
                var continuar = bucleQuery();
                setTimeout(function() {
                    if (continuar != false) {
                        resolve();
                    } else {
                        resultado();
                    }
                    numeroBucle++;
                    console.log(`Bucle numero: ${numeroBucle}`);
                }, tiempo);
            }).then(() => {
                return new Promise(async function(resolve, reject) { //Promesa 4
                    var continuar = bucleQuery();
                    setTimeout(function() {
                        if (continuar != false) {
                            resolve();
                        } else {
                            resultado();
                        }
                        numeroBucle++;
                        console.log(`Bucle numero: ${numeroBucle}`);
                    }, tiempo);
                }).then(() => {
                    return new Promise(async function(resolve, reject) { //Promesa 5
                        var continuar = bucleQuery();
                        setTimeout(function() {
                            if (continuar != false) {
                                resolve();
                            } else {
                                resultado();
                            }
                            numeroBucle++;
                            console.log(`Bucle numero: ${numeroBucle}`);
                        }, tiempo);
                    }).then(() => {
                        return new Promise(async function(resolve, reject) { //Promesa 6
                            var continuar = bucleQuery();
                            setTimeout(function() {
                                if (continuar != false) {
                                    resolve();
                                } else {
                                    resultado();
                                }
                                numeroBucle++;
                                console.log(`Bucle numero: ${numeroBucle}`);
                            }, tiempo);
                        }).then(() => {
                            return new Promise(async function(resolve, reject) { //Promesa 7
                                var continuar = bucleQuery();
                                setTimeout(function() {
                                    if (continuar != false) {
                                        resolve();
                                    } else {
                                        resultado();
                                    }
                                    numeroBucle++;
                                    console.log(`Bucle numero: ${numeroBucle}`);
                                }, tiempo);
                            }).then(() => {
                                return new Promise(async function(resolve, reject) { //Promesa 8
                                    var continuar = bucleQuery();
                                    setTimeout(function() {
                                        if (continuar != false) {
                                            resolve();
                                        } else {
                                            resultado();
                                        }
                                        numeroBucle++;
                                        console.log(`Bucle numero: ${numeroBucle}`);
                                    }, tiempo);
                                }).then(() => {
                                    return new Promise(async function(resolve, reject) { //Promesa 9
                                        var continuar = bucleQuery();
                                        setTimeout(function() {
                                            if (continuar != false) {
                                                resolve();
                                            } else {
                                                resultado();
                                            }
                                            numeroBucle++;
                                            console.log(`Bucle numero: ${numeroBucle}`);
                                        }, tiempo);
                                    }).then(() => {
                                        return new Promise(async function(resolve, reject) { //Promesa 10
                                            var continuar = bucleQuery();
                                            setTimeout(function() {
                                                if (continuar != false) {
                                                    resolve();
                                                } else {
                                                    resultado();
                                                }
                                                numeroBucle++;
                                                console.log(`Bucle numero: ${numeroBucle}`);
                                            }, tiempo);
                                        }).then(() => {
                                            return new Promise(async function(resolve, reject) { //Promesa 11
                                                var continuar = bucleQuery();
                                                setTimeout(function() {
                                                    if (continuar != false) {
                                                        resolve();
                                                    } else {
                                                        resultado();
                                                    }
                                                    numeroBucle++;
                                                    console.log(`Bucle numero: ${numeroBucle}`);
                                                }, tiempo);
                                            }).then(() => {
                                                return new Promise(async function(resolve, reject) { //Promesa 12
                                                    var continuar = bucleQuery();
                                                    setTimeout(function() {
                                                        if (continuar != false) {
                                                            resolve();
                                                        } else {
                                                            resultado();
                                                        }
                                                        numeroBucle++;
                                                        console.log(`Bucle numero: ${numeroBucle}`);
                                                    }, tiempo);
                                                }).then(() => {
                                                    return new Promise(async function(resolve, reject) { //Promesa 13
                                                        var continuar = bucleQuery();
                                                        setTimeout(function() {
                                                            if (continuar != false) {
                                                                resolve();
                                                            } else {
                                                                resultado();
                                                            }
                                                            numeroBucle++;
                                                            console.log(`Bucle numero: ${numeroBucle}`);
                                                        }, tiempo);
                                                    }).then(() => {
                                                        return new Promise(async function(resolve, reject) { //Promesa 14
                                                            var continuar = bucleQuery();
                                                            setTimeout(function() {
                                                                if (continuar != false) {
                                                                    resolve();
                                                                } else {
                                                                    resultado();
                                                                }
                                                                numeroBucle++;
                                                                console.log(`Bucle numero: ${numeroBucle}`);
                                                            }, tiempo);
                                                        }).then(() => {
                                                            return new Promise(async function(resolve, reject) { //Promesa 15
                                                                var continuar = bucleQuery();
                                                                setTimeout(function() {
                                                                    if (continuar != false) {
                                                                        resolve();
                                                                    } else {
                                                                        resultado();
                                                                    }
                                                                    numeroBucle++;
                                                                    console.log(`Bucle numero: ${numeroBucle}`);
                                                                }, tiempo);
                                                            }).then(() => {
                                                                return new Promise(async function(resolve, reject) { //Promesa 16
                                                                    var continuar = bucleQuery();
                                                                    setTimeout(function() {
                                                                        if (continuar != false) {
                                                                            resolve();
                                                                        } else {
                                                                            resultado();
                                                                        }
                                                                        numeroBucle++;
                                                                        console.log(`Bucle numero: ${numeroBucle}`);
                                                                    }, tiempo);
                                                                }).then(() => {
                                                                    return new Promise(async function(resolve, reject) { //Promesa 17
                                                                        var continuar = bucleQuery();
                                                                        setTimeout(function() {
                                                                            if (continuar != false) {
                                                                                resolve();
                                                                            } else {
                                                                                resultado();
                                                                            }
                                                                            numeroBucle++;
                                                                            console.log(`Bucle numero: ${numeroBucle}`);
                                                                        }, tiempo);
                                                                    }).then(() => {
                                                                        return new Promise(async function(resolve, reject) { //Promesa 18
                                                                            var continuar = bucleQuery();
                                                                            setTimeout(function() {
                                                                                if (continuar != false) {
                                                                                    resolve();
                                                                                } else {
                                                                                    resultado();
                                                                                }
                                                                                numeroBucle++;
                                                                                console.log(`Bucle numero: ${numeroBucle}`);
                                                                            }, tiempo);
                                                                        }).then(() => {
                                                                            return new Promise(async function(resolve, reject) { //Promesa 19
                                                                                var continuar = bucleQuery();
                                                                                setTimeout(function() {
                                                                                    if (continuar != false) {
                                                                                        resolve();
                                                                                    } else {
                                                                                        resultado();
                                                                                    }
                                                                                    numeroBucle++;
                                                                                    console.log(`Bucle numero: ${numeroBucle}`);
                                                                                }, tiempo);
                                                                            }).then(() => {
                                                                                resultado();
                                                                            });
                                                                        });
                                                                    });
                                                                });
                                                            });
                                                        });
                                                    });
                                                });
                                            });
                                        });
                                    });
                                });
                            });
                        });
                    });
                });
            });
        });
    });
}

//Función para continuar despues de 1000 rows:
function bucleQuery() {
    if (nextContinuationToken != null) {
        tableSvc.queryEntities(tablaUsar, query, nextContinuationToken, function(error, results, response) {
            if (!error) {
                //Recorrido por row:
                results.entries.forEach(function(entry) {
                    //Logica por row:
                    aceptCount = 0;
                    if (entry['Baja']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['Borrado']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['Check']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['Resguardo']['_'] == "Aprobado") {
                        aceptCount++;
                    }
                    if (entry['HojaDeServicio']['_'] == "Aprobado") {
                        aceptCount++;
                    }

                    console.log(`Aceptados: ${aceptCount}`);
                    if (aceptCount == 5) {
                        //Colocación de la información:
                        //if (entry['Proyecto']['_'] == proyectoTrabajando) {
                        //proyectoCount++;
                        //taskTabla5['PartitionKey']['_'] = entry['Proyecto']['_'];
                        //taskTabla5['RowKey']['_'] = entry['RowKey']['_'];
                        //taskTabla5['Baja']['_'] = entry['Baja']['_'];
                        //taskTabla5['Borrado']['_'] = entry['Borrado']['_'];
                        //taskTabla5['Check']['_'] = entry['Check']['_'];
                        //taskTabla5['Resguardo']['_'] = entry['Resguardo']['_'];
                        //taskTabla5['HolaDeServicio']['_'] = entry['HolaDeServicio']['_'];
                        /*tableSvc.insertEntity(tablaUsar5, taskTabla5, function(error, result, response) {
                            if (!error) {
                                console.log("La entidad se agrego correctamente a la tabla 5");
                            }
                        });*/
                        /*celdaActual++;
                        worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                        worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                        worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                        worksheet.getCell(`D${celdaActual}`).value = entry['Baja']['_'];
                        worksheet.getCell(`E${celdaActual}`).value = entry['Borrado']['_'];
                        worksheet.getCell(`F${celdaActual}`).value = entry['Check']['_'];
                        worksheet.getCell(`G${celdaActual}`).value = entry['Resguardo']['_'];
                        worksheet.getCell(`H${celdaActual}`).value = entry['HojaDeServicio']['_'];*/
                        //}
                        count5++;
                    } else if (aceptCount == 4) {
                        //Colocación de la información:
                        //if (entry['Proyecto']['_'] == proyectoTrabajando) {
                        //proyectoCount++;
                        //taskTabla5['PartitionKey']['_'] = entry['Proyecto']['_'];
                        //taskTabla5['RowKey']['_'] = entry['RowKey']['_'];
                        //taskTabla5['Baja']['_'] = entry['Baja']['_'];
                        //taskTabla5['Borrado']['_'] = entry['Borrado']['_'];
                        //taskTabla5['Check']['_'] = entry['Check']['_'];
                        //taskTabla5['Resguardo']['_'] = entry['Resguardo']['_'];
                        //taskTabla5['HolaDeServicio']['_'] = entry['HolaDeServicio']['_'];
                        /*tableSvc.insertEntity(tablaUsar5, taskTabla5, function(error, result, response) {
                            if (!error) {
                                console.log("La entidad se agrego correctamente a la tabla 5");
                            }
                        });*/
                        /*celdaActual++;
                        worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                        worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                        worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                        worksheet.getCell(`D${celdaActual}`).value = entry['Baja']['_'];
                        worksheet.getCell(`E${celdaActual}`).value = entry['Borrado']['_'];
                        worksheet.getCell(`F${celdaActual}`).value = entry['Check']['_'];
                        worksheet.getCell(`G${celdaActual}`).value = entry['Resguardo']['_'];
                        worksheet.getCell(`H${celdaActual}`).value = entry['HojaDeServicio']['_'];*/
                        //}
                        count4++;
                    } else if (aceptCount == 3) {
                        //Colocación de la información:
                        //if (entry['Proyecto']['_'] == proyectoTrabajando) {
                        //proyectoCount++;
                        //taskTabla5['PartitionKey']['_'] = entry['Proyecto']['_'];
                        //taskTabla5['RowKey']['_'] = entry['RowKey']['_'];
                        //taskTabla5['Baja']['_'] = entry['Baja']['_'];
                        //taskTabla5['Borrado']['_'] = entry['Borrado']['_'];
                        //taskTabla5['Check']['_'] = entry['Check']['_'];
                        //taskTabla5['Resguardo']['_'] = entry['Resguardo']['_'];
                        //taskTabla5['HolaDeServicio']['_'] = entry['HolaDeServicio']['_'];
                        /*tableSvc.insertEntity(tablaUsar5, taskTabla5, function(error, result, response) {
                            if (!error) {
                                console.log("La entidad se agrego correctamente a la tabla 5");
                            }
                        });*/
                        /*celdaActual++;
                        worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                        worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                        worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                        worksheet.getCell(`D${celdaActual}`).value = entry['Baja']['_'];
                        worksheet.getCell(`E${celdaActual}`).value = entry['Borrado']['_'];
                        worksheet.getCell(`F${celdaActual}`).value = entry['Check']['_'];
                        worksheet.getCell(`G${celdaActual}`).value = entry['Resguardo']['_'];
                        worksheet.getCell(`H${celdaActual}`).value = entry['HojaDeServicio']['_'];*/
                        //}
                        count3++;
                    } else if (aceptCount == 2) {
                        //Colocación de la información:
                        //if (entry['Proyecto']['_'] == proyectoTrabajando) {
                        //proyectoCount++;
                        //taskTabla5['PartitionKey']['_'] = entry['Proyecto']['_'];
                        //taskTabla5['RowKey']['_'] = entry['RowKey']['_'];
                        //taskTabla5['Baja']['_'] = entry['Baja']['_'];
                        //taskTabla5['Borrado']['_'] = entry['Borrado']['_'];
                        //taskTabla5['Check']['_'] = entry['Check']['_'];
                        //taskTabla5['Resguardo']['_'] = entry['Resguardo']['_'];
                        //taskTabla5['HolaDeServicio']['_'] = entry['HolaDeServicio']['_'];
                        //tableSvc.insertEntity(tablaUsar5, taskTabla5, function(error, result, response) {
                        //    if (!error) {
                        //        console.log("La entidad se agrego correctamente a la tabla 5");
                        //    }
                        //});
                        /*celdaActual++;
                        worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                        worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                        worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                        worksheet.getCell(`D${celdaActual}`).value = entry['Baja']['_'];
                        worksheet.getCell(`E${celdaActual}`).value = entry['Borrado']['_'];
                        worksheet.getCell(`F${celdaActual}`).value = entry['Check']['_'];
                        worksheet.getCell(`G${celdaActual}`).value = entry['Resguardo']['_'];
                        worksheet.getCell(`H${celdaActual}`).value = entry['HojaDeServicio']['_'];*/
                        //}
                        count2++;
                    } else if (aceptCount == 1) {
                        //Colocación de la información:
                        //if (entry['Proyecto']['_'] == proyectoTrabajando) {
                        //proyectoCount++;
                        //taskTabla5['PartitionKey']['_'] = entry['Proyecto']['_'];
                        //taskTabla5['RowKey']['_'] = entry['RowKey']['_'];
                        //taskTabla5['Baja']['_'] = entry['Baja']['_'];
                        //taskTabla5['Borrado']['_'] = entry['Borrado']['_'];
                        //taskTabla5['Check']['_'] = entry['Check']['_'];
                        //taskTabla5['Resguardo']['_'] = entry['Resguardo']['_'];
                        //taskTabla5['HolaDeServicio']['_'] = entry['HolaDeServicio']['_'];
                        /*tableSvc.insertEntity(tablaUsar5, taskTabla5, function(error, result, response) {
                            if (!error) {
                                console.log("La entidad se agrego correctamente a la tabla 5");
                            }
                        });*/
                        /*celdaActual++;
                        worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                        worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                        worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                        worksheet.getCell(`D${celdaActual}`).value = entry['Baja']['_'];
                        worksheet.getCell(`E${celdaActual}`).value = entry['Borrado']['_'];
                        worksheet.getCell(`F${celdaActual}`).value = entry['Check']['_'];
                        worksheet.getCell(`G${celdaActual}`).value = entry['Resguardo']['_'];
                        worksheet.getCell(`H${celdaActual}`).value = entry['HojaDeServicio']['_'];*/
                        //}
                        count1++;
                    } else {
                        //Colocación de la información:
                        //if (entry['Proyecto']['_'] == proyectoTrabajando) {
                        //proyectoCount++;
                        //taskTabla5['PartitionKey']['_'] = entry['Proyecto']['_'];
                        //taskTabla5['RowKey']['_'] = entry['RowKey']['_'];
                        //taskTabla5['Baja']['_'] = entry['Baja']['_'];
                        //taskTabla5['Borrado']['_'] = entry['Borrado']['_'];
                        //taskTabla5['Check']['_'] = entry['Check']['_'];
                        //taskTabla5['Resguardo']['_'] = entry['Resguardo']['_'];
                        //taskTabla5['HolaDeServicio']['_'] = entry['HolaDeServicio']['_'];
                        /*tableSvc.insertEntity(tablaUsar5, taskTabla5, function(error, result, response) {
                            if (!error) {
                                console.log("La entidad se agrego correctamente a la tabla 5");
                            }
                        });*/
                        /*celdaActual++;
                        worksheet.getCell(`A${celdaActual}`).value = entry['RowKey']['_'];
                        worksheet.getCell(`B${celdaActual}`).value = entry['PartitionKey']['_'];
                        worksheet.getCell(`C${celdaActual}`).value = entry['Proyecto']['_'];
                        worksheet.getCell(`D${celdaActual}`).value = entry['Baja']['_'];
                        worksheet.getCell(`E${celdaActual}`).value = entry['Borrado']['_'];
                        worksheet.getCell(`F${celdaActual}`).value = entry['Check']['_'];
                        worksheet.getCell(`G${celdaActual}`).value = entry['Resguardo']['_'];
                        worksheet.getCell(`H${celdaActual}`).value = entry['HojaDeServicio']['_'];*/
                        //}
                        count0++;
                    }
                });
                //Token que permite continuar despues de leer 1000 rows:
                if (results.continuationToken) {
                    nextContinuationToken = results.continuationToken;
                } else {
                    return false;
                }
            }
        });
    }
}

//Funcion que se ejecuta el final del programa:
function resultado() {
    //taskTabla4['PartitionKey']['_'] = "Proyecto";
    //taskTabla4['RowKey']['_'] = proyectoTrabajando;
    //taskTabla4['NumDoc']['_'] = proyectoCount;
    //taskTabla4['Baja']['_'] = "X";
    //taskTabla4['Borrado']['_'] = "X";
    //taskTabla4['Check']['_'] = "X";
    //taskTabla4['Resguardo']['_'] = "X";
    //taskTabla4['HojaDeServicio']['_'] = "X";

    /*tableSvc.insertEntity(tablaUsar4, taskTabla4, function(error, result, response) {
        if (!error) {
            console.log("La entidad a la tabla 4 se agrego correctamente.");
        }
    });*/
    console.log(`${count5} tienen los 5 campos Aprobados`);
    console.log(`${count4} tienen los 4 campos Aprobados`);
    console.log(`${count3} tienen los 3 campos Aprobados`);
    console.log(`${count2} tienen los 2 campos Aprobados`);
    console.log(`${count1} tienen 1 campo Aprobado`);
    console.log(`${count0} no tienen ningun Aprobado`);
    total = count0 + count1 + count2 + count3 + count4 + count5;
    console.log(`Total de campos analizados: ${total}`);
    console.log(`${proyectoCount} corresponden con el proyecto.`);
}

//Inicia el trabajo:
promesaWork();