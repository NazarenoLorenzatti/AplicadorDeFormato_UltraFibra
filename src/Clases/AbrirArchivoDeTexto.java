package Clases;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import static java.lang.Integer.parseInt;
import static java.lang.String.valueOf;
import java.util.Scanner;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

/*
@author Lorenzatti Nazareno
@version 1.0
nl.loragro@gmail.com
 */

// clase para leer los archivos TXT enviados por pago mis cuentas y link pagos
// para poder imputar los pagos realizados por estos medios.

public class AbrirArchivoDeTexto {

    private String path;

    public AbrirArchivoDeTexto(String path) {

        this.path = path;

    }

    public void leerArchivo(JTable tabla) throws FileNotFoundException {

        Object[] fila = new Object[5];
        int fil = 0;
        int contador = 0;

        DefaultTableModel dt = new DefaultTableModel() {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        // Creamos el header de la tabla
        dt.addColumn("Monto");
        dt.addColumn("Fecha Pago");
        dt.addColumn("Nro DNI Del cliente");
        dt.addColumn("ID cliente");
        dt.addColumn("ID Ref");

        InputStream ins = new FileInputStream(this.path);
        InputStream ins2 = new FileInputStream(this.path);
        
        Scanner scanner = new Scanner(ins);
        Scanner scanner2 = new Scanner(ins2);
        
        // creamos un segundo scanner para contar el total de filas del archivo .TXT
        while(scanner2.hasNextLine()){
            contador++;
            scanner2.nextLine();
        }
        scanner2.close();
        
        // Leemos Linea por linea y vamos creando las filas para la respectiva Jtable        
        // Obteniendo la informacion de acuerdo a la posicion en la que se encuentra
        
        while (scanner.hasNextLine()) {
            String dni = "";
            String Id = "";
            String strFila = scanner.nextLine();

            if (fil != 0 && fil < contador-1) {
                char[] arregloFila = strFila.toCharArray();

                for (int i = 1; i < arregloFila.length; i++) {
                    String posicion = valueOf(arregloFila[i]);

                    if (!posicion.equals(" ")) {
                        dni += posicion;
                    } else {
                        fila[2] = dni;
                        break;
                    }
                }

                for (int i = 20; i < arregloFila.length; i++) {
                    String posicion = valueOf(arregloFila[i]);

                    if (!posicion.equals(" ")) {
                        Id += posicion;
                    } else {
                        fila[3] = Id;
                        fila[4] = Id;
                        break;
                    }
                }
                fila[1] = strFila.substring(49, 57);

                String monto = strFila.substring(58, 68);
                int nroMonto = parseInt(monto);
                fila[0] = String.valueOf(nroMonto);

                dt.addRow(fila);
            }

            fil ++;
        }

        tabla.setModel(dt);
        scanner.close();
    }

}
