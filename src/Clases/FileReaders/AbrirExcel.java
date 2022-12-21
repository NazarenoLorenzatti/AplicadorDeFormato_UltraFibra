package Clases.FileReaders;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Row;

/*
@author Lorenzatti Nazareno
@version 1.0
nl.loragro@gmail.com
*/

// Clase para abrir tabla excel
public class AbrirExcel {

    private String path;
    private FileInputStream inputStream;
    private Workbook wb;
    private Sheet hoja;
    private DataFormatter formatoDeDatos;
    
    public AbrirExcel(String path) {
        
        try {
            this.path = path;
            this.inputStream = new FileInputStream(new File(path));
            this.wb = WorkbookFactory.create(inputStream);
            this.hoja = wb.getSheetAt(0);
            this.formatoDeDatos = new DataFormatter();

        } catch (FileNotFoundException ex) {
            Logger.getLogger(AbrirExcel.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "El fichero no funciona");
        } catch (IOException ex) {
            Logger.getLogger(AbrirExcel.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "No se puede creer el fichero");
        } catch (EncryptedDocumentException ex) {
            Logger.getLogger(AbrirExcel.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Documento encriptado");
        } catch (InvalidFormatException ex) {
            Logger.getLogger(AbrirExcel.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    // Lee la tabla importada y llena el Jtable del formulario
    public void Leer(JTable tabla) {

        DefaultTableModel dt = new DefaultTableModel() {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };

        Iterator i = this.hoja.iterator();
        Object[] fila;
        int fil = 0;

        while (i.hasNext()) {
            Row filaSiguiente = (Row) i.next();
            Iterator iCelda = filaSiguiente.cellIterator();

            // contamos el numero de celdas de la fila
            int largoFila = filaSiguiente.getLastCellNum();
            fila = new Object[largoFila];
            int j = 0;
            if (fil != 0) {
                while (iCelda.hasNext()) {

                    Cell celda = (Cell) iCelda.next();
                    String contenidoCelda = formatoDeDatos.formatCellValue(celda);
                    
                    if (contenidoCelda == null){
                        contenidoCelda = "null";
                    }
                    
                    fila[j] = contenidoCelda;
                    j++;

                }
            } else {
                while (iCelda.hasNext()) { // Si la fila es la numero 0 osea la primera se ponen los encabezados
                    Cell celda = (Cell) iCelda.next();
                    String contenidoCelda = formatoDeDatos.formatCellValue(celda);
                    dt.addColumn(contenidoCelda);                                      
                }

            }
            if (fil != 0){
                dt.addRow(fila);
            } else {
                fil = 1;
            }            
                        
        }
        tabla.setModel(dt);
    }

}
