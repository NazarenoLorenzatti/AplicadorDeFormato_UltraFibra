package Clases.FileWriters;

import java.io.DataOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import javax.swing.JTable;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/*
@author Lorenzatti Nazareno
@version 1.0
nl.loragro@gmail.com
*/

// Clase para exportar el contenido del Jtable a un archivo excel

public class ExportarExcel {

    private final File file;
    private final List tabla;
    private final List nom_files;

    public ExportarExcel(File file, List tabla, List nom_files) throws Exception {
        this.file = file;
        this.tabla = tabla;
        this.nom_files = nom_files;
        if (nom_files.size() != tabla.size()) {
            throw new Exception("Error");
        }
    }

    // Metodo para exportar el contenido del Jtable a un archivo excel, Ultiza Libreria JXl.
    public boolean exportar() throws IOException, WriteException {
        try {
            try (DataOutputStream out = new DataOutputStream(new FileOutputStream(file))) {
                
                WritableWorkbook w = Workbook.createWorkbook(out);
                
                for (int index = 0; index < tabla.size(); index++) {
                    JTable table = (JTable) tabla.get(index);
                    WritableSheet s = w.createSheet((String) nom_files.get(index), 0);
                    
                    for (int i = 0; i < table.getColumnCount(); i++) {
                        String NomCol = table.getColumnName(i);
                        s.addCell(new Label(i, 0, NomCol));
                        
                        for (int j = 0; j < table.getRowCount(); j++) {
                            Object object = table.getValueAt(j, i);
                            s.addCell(new Label(i, j + 1, String.valueOf(object))); //lo pone en la fila 0+1
                        }
                    }
                }
                w.write();
                w.close();
            }
            return true;
        } catch (IOException | WriteException e) {
            System.out.println("Error al exportar a Excel:" + e.getMessage());
            return false;
        }
    }

}
