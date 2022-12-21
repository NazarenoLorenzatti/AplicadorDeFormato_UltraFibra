package Clases.Formatter;

import static java.lang.String.valueOf;
import javax.swing.JTable;

/*
@author Lorenzatti Nazareno
@version 1.0
nl.loragro@gmail.com
*/

//Clase encargada de darle el formato correcto a la informacion para que la tabla
//pueda ser procesada por el banco.

public class FormatoDebitosAutomaticos {

    private final int flag; // Para saber si esta seleccionado debitos abiertos o cerrados. 
    //La informacion se trata diferente si los debitos automaticos son a cuentas del mismo banco (debito cerrado) 
    //o a cuentas de otros bancos (debitos abiertos).
    
    private final String fecha;

    public FormatoDebitosAutomaticos(int flag, String fecha) {

        this.flag = flag;
        this.fecha = fecha;
    }
    
    /*
    Metodo que lee columna por columna y en el caso de que coincida el Header de la columna
    aplica los cambios necesarios para que la informacion se encuentre correcta a la hora 
    de imputar los pagos realizados mediante debito automatico.
    */

    public void normalizar(JTable tabla) {

        for (int j = 0; j < tabla.getColumnCount(); j++) {
            for (int i = 0; i < tabla.getRowCount(); i++) {

                String nombreColumna = tabla.getColumnName(j);

                if (nombreColumna.equals("N°empresa sueldo")) {
                    tabla.setValueAt("00000", i, j);
                }

                // Guarda el tipo de banco el cual son los primeros 3 numeros del CBU en debitos abiertos                
                if (nombreColumna.equals("Cuenta")) {
                    String valorCelda = (String) tabla.getValueAt(i, j);
                    if (this.flag == 1) {
                        String nroSuc = valorCelda.substring(0, 3);
                        tabla.setValueAt(nroSuc, i, j - 3);
                    } else if (this.flag == 2) { // Si son deditos cerrados saca el numero de sucursal del CBU
                        String nroSuc = valorCelda.substring(4, 7);
                        tabla.setValueAt(nroSuc, i, j - 2);
                    }

                }

                if (nombreColumna.equals("Id_adherente") && this.flag == 1) {
                    String celdaNueva = null;
                    String valorCelda = (String) tabla.getValueAt(i, j);
                    String CBU = (String) tabla.getValueAt(i, j-1);
                    String banco = CBU.substring(0, 3);

                    if (banco.equals("072")) { // BANCO SANTANDER
                        if (valorCelda.length() == 4) {
                            celdaNueva = "0" + valorCelda;
                            tabla.setValueAt(celdaNueva, i, j);
                        } else if (valorCelda.length() == 3) {
                            celdaNueva = "00" + valorCelda;
                            tabla.setValueAt(celdaNueva, i, j);
                        } else if (valorCelda.length() == 2) {
                            celdaNueva = "000" + valorCelda;
                            tabla.setValueAt(celdaNueva, i, j);
                        } else if (valorCelda.length() == 1) {
                            celdaNueva = "0000" + valorCelda;
                            tabla.setValueAt(celdaNueva, i, j);
                        }
                      
                    }

                    if (banco.equals("007")) { // BANCO GALICIA
                        if (valorCelda.length() == 5) {
                            celdaNueva = "0" + valorCelda;
                            tabla.setValueAt(celdaNueva, i, j);
                        } else if (valorCelda.length() == 4) {
                            celdaNueva = "00" + valorCelda;
                            tabla.setValueAt(celdaNueva, i, j);
                        } else if (valorCelda.length() == 3) {
                            celdaNueva = "000" + valorCelda;
                            tabla.setValueAt(celdaNueva, i, j);
                        } else if (valorCelda.length() == 2) {
                            celdaNueva = "0000" + valorCelda;
                            tabla.setValueAt(celdaNueva, i, j);
                        } else if (valorCelda.length() == 1) {
                            celdaNueva = "00000" + valorCelda;
                            tabla.setValueAt(celdaNueva, i, j);
                        }
                         
                    }
                    
                }

                // Al tipo de Banco de debitos cerrados se le coloca a todos el fijo 285 ya que tambien admite Cuentas corrientes.
                if (nombreColumna.equals("Tipo_Banco") && this.flag == 2) {
                    tabla.setValueAt("285", i, j);
                }

                // La Columna Sucursal de los debitos abiertos siempre se completa con 0100
                if (nombreColumna.equals("Sucursal") && this.flag == 1) {
                    tabla.setValueAt("0100", i, j);
                }

                if (nombreColumna.equals("Tipo_cuenta")) {

                    String valorCelda = (String) tabla.getValueAt(i, j);

                    // si son debitos abiertos
                    if (this.flag == 1) {
                        tabla.setValueAt(0, i, j);

                        // Si son debitos cerrados
                    } else if (this.flag == 2) {

                        String tipoCuenta = valorCelda.substring(0, 0);

                        if (tipoCuenta.equals("3")) {
                            tabla.setValueAt(3, i, j);
                        } else if (tipoCuenta.equals("4")) {
                            tabla.setValueAt(4, i, j);
                        } else {
                            tabla.setValueAt(0, i, j);
                        }

                    }

                }

                if (nombreColumna.equals("Fecha_vto")) {
                    tabla.setValueAt(this.fecha, i, j);
                }

                if (nombreColumna.equals("Moneda")) {
                    tabla.setValueAt("080", i, j);
                }

                // Normaliza el importe segun el formato pedido por el banco, teniendo en cuenta los decimale
                if (nombreColumna.equals("Importe")) {
                    String celdaNueva = null;
                    String valorCelda = (String) tabla.getValueAt(i, j);
                    char[] strArr = valorCelda.toCharArray();
                    boolean flag2 = true;

                    for (int index = 0; index < strArr.length; index++) {

                        String strCar = valueOf(strArr[index]);

                        if (strCar.equals(",")) {
                            
                            String subString = valorCelda.substring(index);
                            
                            if(subString.length() < 3 ){
                                
                                celdaNueva = valorCelda.replace(",", "");
                                celdaNueva = celdaNueva + "0";
                            }else {
                               celdaNueva = valorCelda.replace(",", ""); 
                            }
                            
                            flag2 = false;
                        }

                    }

                    if (flag2) {
                        celdaNueva = valorCelda + "00";
                    }

                    tabla.setValueAt(celdaNueva, i, j);
                }

            }
        }

    }
}
