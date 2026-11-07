// config.js - Cámaras CIO Esval - Adv
// Puedes pegar el ID del video o el enlace completo de YouTube (watch?v= o live/)
const camarasConfig = [
    { 
        nombre: "Valparaíso", 
        tipo: "youtube", 
        enlace: "https://www.youtube.com/watch?v=5jPpMkg5daM" 
    },
    { 
        nombre: "El Quisco", 
        tipo: "youtube", 
        enlace: "https://www.youtube.com/watch?v=8mGoV-mbwTk" 
    },
    { 
        nombre: "Cartagenas", 
        tipo: "youtube", 
        enlace: "https://www.youtube.com/live/obYSVVM7RB0?si=DsCvNqaasPZwZmX8" 
    },
    { 
        nombre: "La Serena", 
        tipo: "youtube", 
        enlace: "https://www.youtube.com/watch?v=S-yZwnKlMsY" 
    },
    { 
        nombre: "Viña del Mar", 
        tipo: "youtube", 
        enlace: "https://www.youtube.com/live/mxirpLYPG3U?si=axvvhO7sqew6s0Ho" 
    }
];




function onEdit(e) {
  if (!e) return;
  
  var rangoEditado = e.range;
  var hojaIngreso = rangoEditado.getSheet();
  
  if (hojaIngreso.getName() === "Calculadora" && rangoEditado.getA1Notation() === "C2") {
    
    if (rangoEditado.getValue() === true) {
      
      var item = hojaIngreso.getRange("A2").getValue();
      var valor = hojaIngreso.getRange("B2").getValue();
      
      var observacion = hojaIngreso.getRange("A3").getValue();
      
      if (item === "" || valor === "") {
        rangoEditado.setValue(false); 
        return;
      }
      
      var meses = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"];
      var fecha = new Date();
      
      var mesActual = fecha.getMonth();
      var diaActual = fecha.getDate();
      
      // NUEVO: Si el día es mayor a 21 (es decir, del 22 en adelante), suma 1 al mes
      var indiceMes = (diaActual > 21) ? (mesActual + 1) % 12 : mesActual;
      var nombreMes = meses[indiceMes]; 
      
      var hojaDestino = e.source.getSheetByName(nombreMes);
      
      if (!hojaDestino) {
        rangoEditado.setValue(false);
        return;
      }
      
      var valoresColumna = hojaDestino.getRange("E:E").getValues();
      var ultimaFila = 0;
      
      for (var i = valoresColumna.length - 1; i >= 0; i--) {
        if (valoresColumna[i][0] !== "") {
          ultimaFila = i + 1; 
          break;
        }
      }
      
      var filaDestino = ultimaFila + 1; 
      
      hojaDestino.getRange(filaDestino, 5, 1, 4).setValues([[item, valor, fecha, observacion]]);
      
      hojaDestino.getRange(filaDestino, 7).setNumberFormat("dd/MM/yyyy HH:mm:ss");
      
      hojaIngreso.getRange("A2").clearContent();
      hojaIngreso.getRange("B2").clearContent();
      hojaIngreso.getRange("A3").clearContent();
      
      rangoEditado.setValue(false);
    }
  }
}
