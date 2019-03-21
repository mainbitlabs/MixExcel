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
//Contadores:
var contador = 0;
var contadorRepeticion = 0;
var contadorAprobados = 0;
var iguales = 0;
var hojadeserviciocount = 0;
var bajacount = 0;
var resguardocount = 0;
var checkcount = 0;
var borradocount = 0;
var contadorDeCaracteres = 0;
var contadorX = 0;
var nombreDelProyectoExcel = "";
//Contenedores:
var baja = "";
var borrado = "";
var check = "";
var hojadeservicio = "";
var resguardo = "";
var proyecto = "";
//Excel:
var celdaActual3Aprobados = 1;
var celdaActualNoCumple = 1;
//Control de do-while:
var finalizar = false;
//Variables del proyecto:
var proyectoTrabajando = "INGRAM -IMPLEMENTACION";
var bajaExiste = "";
var borradoExiste = "";
var checkExiste = "X";
var resguardoExiste = "X";
var hojaDeServicioExiste = "";

//Programa
async function working() {

    //Contador de X:
    if (bajaExiste == "X") {
        contadorX++;
    }
    if (borradoExiste == "X") {
        contadorX++;
    }
    if (checkExiste == "X") {
        contadorX++;
    }
    if (resguardoExiste == "X") {
        contadorX++;
    }
    if (hojaDeServicioExiste == "X") {
        contadorX++;
    }

    //Contar caracteres del proyecto.
    contadorDeCaracteres = proyectoTrabajando.length;

    //Colocación del titulo de cada columna del libro 3Aprobados:
    worksheet3Aprobados.getCell(`A${celdaActual3Aprobados}`).value = 'RowKey';
    worksheet3Aprobados.getCell(`B${celdaActual3Aprobados}`).value = 'Borrado';
    worksheet3Aprobados.getCell(`C${celdaActual3Aprobados}`).value = 'Check';
    worksheet3Aprobados.getCell(`D${celdaActual3Aprobados}`).value = 'Resguardo';
    worksheet3Aprobados.getCell(`E${celdaActual3Aprobados}`).value = 'Baja';
    worksheet3Aprobados.getCell(`F${celdaActual3Aprobados}`).value = 'HojaDeServicio';
    worksheet3Aprobados.getCell(`G${celdaActual3Aprobados}`).value = 'Proyecto';

    //Colocación del titulo de cada columna del libro NoCumple:
    worksheetNoCumple.getCell(`A${celdaActualNoCumple}`).value = 'RowKey';
    worksheetNoCumple.getCell(`B${celdaActualNoCumple}`).value = 'Borrado';
    worksheetNoCumple.getCell(`C${celdaActualNoCumple}`).value = 'Check';
    worksheetNoCumple.getCell(`D${celdaActualNoCumple}`).value = 'Resguardo';
    worksheetNoCumple.getCell(`E${celdaActualNoCumple}`).value = 'Baja';
    worksheetNoCumple.getCell(`F${celdaActualNoCumple}`).value = 'HojaDeServicio';
    worksheetNoCumple.getCell(`G${celdaActualNoCumple}`).value = 'Proyecto';

    //Aumento de celdas Excel:
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

                    //Contador de entidades encontradas
                    contador++;

                    //Reiniciar contadores:
                    borradocount = 0;
                    bajacount = 0;
                    checkcount = 0;
                    resguardocount = 0;
                    hojadeserviciocount = 0;
                    contadorRepeticion = 0;
                    contadorAprobados = 0;

                    //Vaciar contenedores_
                    borrado = "";
                    baja = "";
                    check = "";
                    resguardo = "";
                    hojadeservicio = "";
                    proyecto = "";

                    console.log(`${entry['RowKey']['_']}`);

                    for (var key in dataAzul) { //Inicia comparación desde libro.
                        if (entry['RowKey']['_'] == dataAzul[key]['Serie']) { // Detectar concidencias con la entidad actual.

                            //Extraer y cortar el nombre de proyecto desde Excel:
                            nombreDelProyectoExcel = dataAzul[key]['Nombre'];
                            nombreDelProyectoExcel = nombreDelProyectoExcel.slice(0, contadorDeCaracteres);

                            //Detectar el tipo de documento seleccionado:
                            if (dataAzul[key]['TipoDoc'] == "Borrado" && borradoExiste == "X") {
                                borrado = dataAzul[key]['Estatus'];
                                if (dataAzul[key]['Estatus'] == "Aprobado") {
                                    contadorAprobados++;
                                }
                                borradocount++;
                            } else if (dataAzul[key]['TipoDoc'] == "Check" && checkExiste == "X") {
                                check = dataAzul[key]['Estatus'];
                                if (dataAzul[key]['Estatus'] == "Aprobado") {
                                    contadorAprobados++;
                                }
                                checkcount++;
                            } else if (dataAzul[key]['TipoDoc'] == "Resguardo" && resguardoExiste == "X") {
                                resguardo = dataAzul[key]['Estatus'];
                                if (dataAzul[key]['Estatus'] == "Aprobado") {
                                    contadorAprobados++;
                                }
                                resguardocount++;
                            } else if (dataAzul[key]['TipoDoc'] == "HojaDeServicio" && hojaDeServicioExiste == "X") {
                                hojadeservicio = dataAzul[key]['Estatus'];
                                if (dataAzul[key]['Estatus'] == "Aprobado") {
                                    contadorAprobados++;
                                }
                                hojadeserviciocount++;
                            } else if (dataAzul[key]['TipoDoc'] == "Baja" && bajaExiste == "X") {
                                baja = dataAzul[key]['Estatus'];
                                if (dataAzul[key]['Estatus'] == "Aprobado") {
                                    contadorAprobados++;
                                }
                                bajacount++;
                            }

                            console.log(`Son iguales: ${entry['RowKey']['_']} - ${dataAzul[key]['Serie']}`);
                        }
                        //Contenedor del nombre de proyecto:
                        proyecto = nombreDelProyectoExcel;
                    }

                    if (contadorAprobados == contadorX && proyectoTrabajando == proyecto) {
                        //Escribir en el libro Aprobados:
                        worksheet3Aprobados.getCell(`A${celdaActual3Aprobados}`).value = entry['RowKey']['_'];
                        worksheet3Aprobados.getCell(`B${celdaActual3Aprobados}`).value = borrado;
                        worksheet3Aprobados.getCell(`C${celdaActual3Aprobados}`).value = check;
                        worksheet3Aprobados.getCell(`D${celdaActual3Aprobados}`).value = resguardo;
                        worksheet3Aprobados.getCell(`E${celdaActual3Aprobados}`).value = hojadeservicio;
                        worksheet3Aprobados.getCell(`F${celdaActual3Aprobados}`).value = baja;
                        worksheet3Aprobados.getCell(`G${celdaActual3Aprobados}`).value = proyecto;
                        celdaActual3Aprobados++;
                        iguales++;
                    } else if (proyectoTrabajando == proyecto) {
                        if (contadorRepeticion <= 1) {
                            //Colocación del titulo de cada columna del libro NoCumple:
                            worksheetNoCumple.getCell(`A${celdaActualNoCumple}`).value = entry['RowKey']['_'];
                            worksheetNoCumple.getCell(`B${celdaActualNoCumple}`).value = borrado;
                            worksheetNoCumple.getCell(`C${celdaActualNoCumple}`).value = check;
                            worksheetNoCumple.getCell(`D${celdaActualNoCumple}`).value = resguardo;
                            worksheetNoCumple.getCell(`E${celdaActualNoCumple}`).value = hojadeservicio;
                            worksheetNoCumple.getCell(`F${celdaActualNoCumple}`).value = baja;
                            worksheetNoCumple.getCell(`G${celdaActualNoCumple}`).value = proyecto;
                            contadorRepeticion++;
                        }
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
        workbook3Aprobados.xlsx.writeFile('Aprobados.xlsx').then(function() { //Puedes colocar cualquier nombre al archivo final sustituyendo "final.xlsx" (recuerda respetar siempre la extención .xlsx).
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

    //Resumen por consola:
    console.log(`Se encontrarion ${contador} entidades.`);
    console.log(`Se encontrarion ${iguales} que coinciden con el trabajo.`);
    console.log("El programa ha terminado:");
}

//Inicia el trabajo:
working();