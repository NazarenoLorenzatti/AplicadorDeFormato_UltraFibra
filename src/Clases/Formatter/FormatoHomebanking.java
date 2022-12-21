package Clases.Formatter;

import static java.lang.String.valueOf;
import javax.swing.JTable;

/*
@author Lorenzatti Nazareno
@version 1.0
nl.loragro@gmail.com
 */

// Clase encargada de convertir la informacion de los archivos de texto de cobranza enviados por pago mis cuentas y link pagos
// a una Jtable formateandola para que pueda imputarse de forma masiva. Luego el Jtable se exporta a un excel.
public class FormatoHomebanking {

    public FormatoHomebanking() {

    }

    /*
    Metodo que lee columna por columna y en el caso de que coincida el Header de la columna
    aplica los cambios necesarios para que la informacion se encuentre correcta a la hora 
    de imputar los pagos realizados mediante links de pago.
     */
    public void normalizar(JTable tabla) {

        for (int j = 0; j < tabla.getColumnCount(); j++) {
            for (int i = 0; i < tabla.getRowCount(); i++) {

                String nombreColumna = tabla.getColumnName(j);
                
                
                 if (nombreColumna.equals("Monto")) {
                    String valorCelda = (String) tabla.getValueAt(i, j);
                    
                    String decimales = valorCelda.substring(valorCelda.length()-2);
                    String enteros =  valorCelda.substring( 0, valorCelda.length()-2);
                    
                    valorCelda = enteros + "," + decimales;
                    tabla.setValueAt(valorCelda, i, j);

                }

                // En caso de que el dia y el mes comiencen con 0 los quita ya que asi requiere el formato
                if (nombreColumna.equals("Fecha Pago")) {
                    String valorCelda = (String) tabla.getValueAt(i, j);

                    
                    String año = valorCelda.substring(0, 4);
                    String mes = valorCelda.substring(4, 6);
                    String dia = valorCelda.substring(6, 8);

                    char car = dia.charAt(0);
                    String comp = valueOf(car);

                    if (comp.equals("0")) {
                        dia = dia.replaceFirst("0", "");
                    }

                    car = mes.charAt(0);
                    comp = valueOf(car);
                    if (comp.equals("0")) {
                        mes = mes.replaceFirst("0", "");

                    }
                    String fecha = dia +"/"+ mes+"/"+ año;
                    tabla.setValueAt(fecha, i, j);

                }
            }
        }
    }
}
