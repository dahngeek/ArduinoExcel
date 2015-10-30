//ArduinoExcel
//  -- Lee datos del serial, que envia una arduino con un sensor TMP36... 
//  -- los tabula en un documento de excel.. y los guarda en una base de datos de MongoDB

// Autor: Daniel Xutuc (dahngeek.com)

sp = require("serialport");
SerialPort = sp.SerialPort;
var excelbuilder = require('msexcel-builder');
var keypress = require('keypress');

// make `process.stdin` begin emitting "keypress" events 
keypress(process.stdin);



var portName = 'COM41'; //Puerto serial del arduino

function recibeSerial(debug) {
	var receivedData = "";
    serialPort = new SerialPort(portName, {
        baudrate: 9600, //velocidad
        // defaults para arduino...
        parser: sp.parsers.readline("\n")
    });
    serialPort.on('open', function(){
    	console.log('Se estableció conexión serial.');
    	var infoData = [];
    	infoData.push({equis: "x", val:"y"});
    	//Documento de excel
        var i = 0;
    	var workbook = excelbuilder.createWorkbook('./', 'datosAnalisisNombredeArchivo.xlsx');
    	  //var sheet1 = workbook.createSheet('Prueba2', 10, 12); //Hoja de excel         
        // listen for the "keypress" event 
        process.stdin.on('keypress', function (ch, key) {
          console.log('got "keypress"', key);
          if (key && key.ctrl && key.name == 'u') {
                    //Cerrar puerto Serial
                    serialPort.close();
                    //Crear hoja en el documento de excel
                    var sheet1 = workbook.createSheet('test', 3, infoData.length);
                    

                    //Registrar fila con dato
                    function registraColumna(elem, index, arr) {
                        console.log("registrnado row: "+elem.equis+" | "+elem.val+" index:"+index);
                        sheet1.set(1, index+1, elem.equis);
                        sheet1.set(2, index+1, elem.val);
                    }
                    infoData.forEach(registraColumna);
                    //Guardar
                    workbook.save(function(ok){
                        if (!ok) 
                          workbook.cancel();
                        else
                          console.log('congratulations, your workbook created');
                      });
          }
        });

    	serialPort.on('data', function(data) {
                receivedData += data.toString();
                
    		if (receivedData.length !== 0 || receivedData !="") {
                //console.log(receivedData);
                
    			var valorY = receivedData;
    			var valorX = i =i+1;
                
                //Guardarlo en la base de datos de Mongo
  			
                    infoData.push({equis: valorX, val: valorY});
                    receivedData = "";
                }
    			
    		}
    	);
    });
}

//Inizializar :D
recibeSerial();