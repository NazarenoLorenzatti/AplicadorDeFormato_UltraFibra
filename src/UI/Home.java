package UI;

import Clases.FileReaders.AbrirArchivoDeTexto;
import Clases.FileReaders.AbrirExcel;
import Clases.FileWriters.ExportarExcel;
import Clases.Formatter.FormatoDebitosAutomaticos;
import Clases.Formatter.FormatoHomebanking;
import Clases.Formatter.FormatoMacroClicks;
import java.awt.Color;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Image;
import java.awt.Toolkit;
import java.io.File;
import java.io.FileNotFoundException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import static javax.swing.JOptionPane.YES_NO_OPTION;
import javax.swing.JPanel;
import javax.swing.JTable;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;

/*
@author Lorenzatti Nazareno
@version 1.0
nl.loragro@gmail.com
 */
// Formulario para el manejo del normalizador
public final class Home extends javax.swing.JFrame {

    private FondoPanel fondo = new FondoPanel();
    private AbrirExcel abrirEx;
    private AbrirArchivoDeTexto abrirTxt;
    private String file;
    private ExportarExcel e;

    public Home() {
        this.setContentPane(this.fondo);
        initComponents();
        jBtnExportar.setEnabled(false);
        this.setLocationRelativeTo(null);

        ImageIcon imgIcon = (ImageIcon) jLabelLogo.getIcon();
        Image imgEscalada = imgIcon.getImage().getScaledInstance(jLabelLogo.getWidth(), jLabelLogo.getHeight(), Image.SCALE_SMOOTH);
        Icon iconoEscalado = new ImageIcon(imgEscalada);

        jLabelLogo.setIcon(iconoEscalado);

        if (rBtnAbiertos.isSelected()) {
            rBtnAbiertos.setBorderPainted(true);
            rBtnAbiertos.setFont(new Font("Segoe UI", 1, 14));
            rBtnCerrados.setContentAreaFilled(false);
        }
    }

    @Override
    public Image getIconImage() {
        Image retValue = Toolkit.getDefaultToolkit().getImage(ClassLoader.getSystemResource("Imagenes/Icono.png"));
        return retValue;
    }

    // Metodo para importar la tabla excel a un Jtable
    public void importar(JTable tabla) {
        String directorio = null;
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos Excel", "xls", "xlsx");
        fc.setFileFilter(filter);

        int seleccion = fc.showOpenDialog(this);

        if (seleccion == JFileChooser.APPROVE_OPTION) {

            File fichero = fc.getSelectedFile();
            directorio = fichero.getAbsolutePath();

            abrirEx = new AbrirExcel(directorio);
            abrirEx.Leer(tabla);

        }
    }

    // metodo para Importar el Jtable a un archivo TXT
    public void importarTxt(JTable tabla) throws FileNotFoundException {
        String directorio = null;
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos TXT", "txt");
        fc.setFileFilter(filter);

        int seleccion = fc.showOpenDialog(this);

        if (seleccion == JFileChooser.APPROVE_OPTION) {

            File fichero = fc.getSelectedFile();
            directorio = fichero.getAbsolutePath();

            abrirTxt = new AbrirArchivoDeTexto(directorio);
            abrirTxt.leerArchivo(tabla);

        }
    }

    // metodo para exportar el Jtable a un archivo excel
    public void exportar(JTable tabla) {
        if (tabla.getRowCount() > 0) {
            JFileChooser chooser = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos de excel", "xls", "xlsx");
            chooser.setFileFilter(filter);
            chooser.setDialogTitle("Guardar archivo");
            chooser.setAcceptAllFileFilterUsed(false);
            if (chooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION) {
                List tb = new ArrayList();
                List nom = new ArrayList();
                tb.add(tabla);
                nom.add("Normalizador");
                this.file = chooser.getSelectedFile().toString().concat(".xls");
                try {
                    this.e = new ExportarExcel(new File(this.file), tb, nom);
                    if (e.exportar()) {
                        JOptionPane.showMessageDialog(null, "Los datos fueron exportados a excel en el directorio seleccionado", "Mensaje de Informacion", JOptionPane.INFORMATION_MESSAGE);
                        File path = new File(this.file);
                        Desktop.getDesktop().open(path);
                    }
                } catch (Exception e) {
                    JOptionPane.showMessageDialog(null, "Hubo un error " + e.getMessage(), " Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        } else {
            JOptionPane.showMessageDialog(this, "No hay datos para exportar", "Mensaje de error", JOptionPane.ERROR_MESSAGE);
        }
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        buttonGroup1 = new javax.swing.ButtonGroup();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        jPaneDebitos = new FondoPanel();
        JbtnImportar = new javax.swing.JButton();
        jScrollPane2 = new javax.swing.JScrollPane();
        jTableDatosExcel2 = new javax.swing.JTable();
        jBtnNormalizador = new javax.swing.JButton();
        jBtnExportar = new javax.swing.JButton();
        rBtnAbiertos = new javax.swing.JRadioButton();
        rBtnCerrados = new javax.swing.JRadioButton();
        jdc_fecha = new com.toedter.calendar.JDateChooser();
        jLabelLogo = new javax.swing.JLabel();
        jLabel1 = new javax.swing.JLabel();
        jPanelLinkPagos = new FondoPanel();
        JbtnImportarLinks = new javax.swing.JButton();
        jBtnNormalizadorLinks = new javax.swing.JButton();
        jBtnExportarLinks = new javax.swing.JButton();
        jScrollPane3 = new javax.swing.JScrollPane();
        tablaLinks = new javax.swing.JTable();
        jLabel2 = new javax.swing.JLabel();
        jPanelLinkPagos1 = new FondoPanel();
        JbtnImportarPMC = new javax.swing.JButton();
        jBtnNormalizadorPMC = new javax.swing.JButton();
        jBtnExportarPMC = new javax.swing.JButton();
        jScrollPane4 = new javax.swing.JScrollPane();
        tablaPMC = new javax.swing.JTable();
        jLabel3 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Normalizador Planillas Debitos");
        setIconImage(getIconImage());
        setResizable(false);

        jTabbedPane1.setBorder(javax.swing.BorderFactory.createMatteBorder(3, 3, 3, 3, new java.awt.Color(102, 102, 102)));
        jTabbedPane1.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jTabbedPane1.setFont(new java.awt.Font("Segoe UI", 1, 18)); // NOI18N
        jTabbedPane1.setName(""); // NOI18N

        jPaneDebitos.setBackground(new java.awt.Color(0, 0, 102));
        jPaneDebitos.setOpaque(false);

        JbtnImportar.setBackground(new java.awt.Color(247, 247, 247));
        JbtnImportar.setFont(new java.awt.Font("Segoe UI", 3, 18)); // NOI18N
        JbtnImportar.setForeground(new java.awt.Color(0, 0, 0));
        JbtnImportar.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/importar.png"))); // NOI18N
        JbtnImportar.setText("Importar Excel");
        JbtnImportar.setBorder(javax.swing.BorderFactory.createCompoundBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255), 3), javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED)));
        JbtnImportar.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        JbtnImportar.setFocusPainted(false);
        JbtnImportar.setPressedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Importar Pressed.png"))); // NOI18N
        JbtnImportar.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/importar roll.png"))); // NOI18N
        JbtnImportar.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                JbtnImportarMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                JbtnImportarMouseExited(evt);
            }
        });
        JbtnImportar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                JbtnImportarActionPerformed(evt);
            }
        });

        jScrollPane2.setOpaque(false);

        jTableDatosExcel2.setFont(jTableDatosExcel2.getFont());
        jTableDatosExcel2.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {

            }
        ));
        jTableDatosExcel2.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_ALL_COLUMNS);
        jTableDatosExcel2.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jTableDatosExcel2.setOpaque(false);
        jScrollPane2.setViewportView(jTableDatosExcel2);

        jBtnNormalizador.setBackground(new java.awt.Color(247, 247, 247));
        jBtnNormalizador.setFont(new java.awt.Font("Segoe UI", 3, 18)); // NOI18N
        jBtnNormalizador.setForeground(new java.awt.Color(0, 0, 0));
        jBtnNormalizador.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Normalizar.png"))); // NOI18N
        jBtnNormalizador.setText("Normalizador");
        jBtnNormalizador.setBorder(javax.swing.BorderFactory.createCompoundBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255), 3), javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED)));
        jBtnNormalizador.setBorderPainted(false);
        jBtnNormalizador.setContentAreaFilled(false);
        jBtnNormalizador.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jBtnNormalizador.setEnabled(false);
        jBtnNormalizador.setFocusPainted(false);
        jBtnNormalizador.setPressedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Normalizar Pressed.png"))); // NOI18N
        jBtnNormalizador.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/normalizar Roll.png"))); // NOI18N
        jBtnNormalizador.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jBtnNormalizadorMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jBtnNormalizadorMouseExited(evt);
            }
        });
        jBtnNormalizador.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBtnNormalizadorActionPerformed(evt);
            }
        });

        jBtnExportar.setBackground(new java.awt.Color(247, 247, 247));
        jBtnExportar.setFont(new java.awt.Font("Segoe UI", 3, 18)); // NOI18N
        jBtnExportar.setForeground(new java.awt.Color(0, 0, 0));
        jBtnExportar.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Exportar.png"))); // NOI18N
        jBtnExportar.setText("Exportar Excel");
        jBtnExportar.setBorder(javax.swing.BorderFactory.createCompoundBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255), 3), javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED)));
        jBtnExportar.setBorderPainted(false);
        jBtnExportar.setContentAreaFilled(false);
        jBtnExportar.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jBtnExportar.setFocusPainted(false);
        jBtnExportar.setPressedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Exportar Pressed.png"))); // NOI18N
        jBtnExportar.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Exportar Roll.png"))); // NOI18N
        jBtnExportar.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jBtnExportarMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jBtnExportarMouseExited(evt);
            }
        });
        jBtnExportar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBtnExportarActionPerformed(evt);
            }
        });

        rBtnAbiertos.setBackground(new java.awt.Color(102, 156, 187));
        buttonGroup1.add(rBtnAbiertos);
        rBtnAbiertos.setFont(new java.awt.Font("Segoe UI", 3, 14)); // NOI18N
        rBtnAbiertos.setForeground(new java.awt.Color(0, 0, 0));
        rBtnAbiertos.setText("Debitos Abiertos");
        rBtnAbiertos.setBorder(javax.swing.BorderFactory.createMatteBorder(3, 3, 3, 3, new java.awt.Color(102, 102, 102)));
        rBtnAbiertos.setContentAreaFilled(false);
        rBtnAbiertos.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        rBtnAbiertos.setFocusPainted(false);
        rBtnAbiertos.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        rBtnAbiertos.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        rBtnAbiertos.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Abiertos.png"))); // NOI18N
        rBtnAbiertos.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Abiertos Roll.png"))); // NOI18N
        rBtnAbiertos.setSelectedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Abiertos Selected.png"))); // NOI18N
        rBtnAbiertos.setVerticalAlignment(javax.swing.SwingConstants.TOP);
        rBtnAbiertos.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        rBtnAbiertos.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                rBtnAbiertosMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                rBtnAbiertosMouseExited(evt);
            }
        });
        rBtnAbiertos.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                rBtnAbiertosActionPerformed(evt);
            }
        });

        rBtnCerrados.setBackground(new java.awt.Color(247, 247, 253));
        buttonGroup1.add(rBtnCerrados);
        rBtnCerrados.setFont(new java.awt.Font("Segoe UI", 3, 14)); // NOI18N
        rBtnCerrados.setForeground(new java.awt.Color(0, 0, 0));
        rBtnCerrados.setText("Debitos Cerrados");
        rBtnCerrados.setBorder(javax.swing.BorderFactory.createMatteBorder(3, 3, 3, 3, new java.awt.Color(102, 102, 102)));
        rBtnCerrados.setContentAreaFilled(false);
        rBtnCerrados.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        rBtnCerrados.setFocusPainted(false);
        rBtnCerrados.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        rBtnCerrados.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        rBtnCerrados.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Cerrados.png"))); // NOI18N
        rBtnCerrados.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Cerrados Roll.png"))); // NOI18N
        rBtnCerrados.setSelectedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Cerrados Selected.png"))); // NOI18N
        rBtnCerrados.setVerticalAlignment(javax.swing.SwingConstants.TOP);
        rBtnCerrados.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        rBtnCerrados.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                rBtnCerradosMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                rBtnCerradosMouseExited(evt);
            }
        });
        rBtnCerrados.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                rBtnCerradosActionPerformed(evt);
            }
        });

        jdc_fecha.setBackground(new java.awt.Color(0, 0, 102));
        jdc_fecha.setForeground(new java.awt.Color(0, 0, 0));
        jdc_fecha.setDateFormatString("dd/MM/yyyy");
        jdc_fecha.setFont(new java.awt.Font("Segoe UI Historic", 1, 14)); // NOI18N
        jdc_fecha.setOpaque(false);

        jLabelLogo.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/UltraFibra.png"))); // NOI18N
        jLabelLogo.setText("jLabel2");

        jLabel1.setFont(new java.awt.Font("Segoe UI", 3, 14)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(0, 0, 0));
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("Seleccione Fecha de Vencimiento");
        jLabel1.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);

        javax.swing.GroupLayout jPaneDebitosLayout = new javax.swing.GroupLayout(jPaneDebitos);
        jPaneDebitos.setLayout(jPaneDebitosLayout);
        jPaneDebitosLayout.setHorizontalGroup(
            jPaneDebitosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPaneDebitosLayout.createSequentialGroup()
                .addGroup(jPaneDebitosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPaneDebitosLayout.createSequentialGroup()
                        .addGap(71, 71, 71)
                        .addComponent(JbtnImportar, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(80, 80, 80)
                        .addComponent(jBtnNormalizador, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(88, 88, 88)
                        .addComponent(jBtnExportar, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPaneDebitosLayout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 1040, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPaneDebitosLayout.createSequentialGroup()
                        .addGap(132, 132, 132)
                        .addGroup(jPaneDebitosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jdc_fecha, javax.swing.GroupLayout.DEFAULT_SIZE, 272, Short.MAX_VALUE)
                            .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addGap(60, 60, 60)
                        .addComponent(rBtnCerrados, javax.swing.GroupLayout.PREFERRED_SIZE, 134, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(58, 58, 58)
                        .addComponent(rBtnAbiertos, javax.swing.GroupLayout.PREFERRED_SIZE, 137, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPaneDebitosLayout.createSequentialGroup()
                        .addGap(234, 234, 234)
                        .addComponent(jLabelLogo, javax.swing.GroupLayout.PREFERRED_SIZE, 546, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap(24, Short.MAX_VALUE))
        );
        jPaneDebitosLayout.setVerticalGroup(
            jPaneDebitosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPaneDebitosLayout.createSequentialGroup()
                .addGap(14, 14, 14)
                .addComponent(jLabelLogo)
                .addGap(34, 34, 34)
                .addGroup(jPaneDebitosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(rBtnCerrados, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(rBtnAbiertos)
                    .addGroup(jPaneDebitosLayout.createSequentialGroup()
                        .addComponent(jdc_fecha, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 58, Short.MAX_VALUE)
                .addGroup(jPaneDebitosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(JbtnImportar, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jBtnNormalizador, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jBtnExportar, javax.swing.GroupLayout.PREFERRED_SIZE, 83, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(28, 28, 28)
                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 366, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(35, 35, 35))
        );

        jTabbedPane1.addTab("NORMALIZADOR DEBITOS", jPaneDebitos);

        JbtnImportarLinks.setBackground(new java.awt.Color(247, 247, 247));
        JbtnImportarLinks.setFont(new java.awt.Font("Segoe UI", 3, 18)); // NOI18N
        JbtnImportarLinks.setForeground(new java.awt.Color(0, 0, 0));
        JbtnImportarLinks.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/importar.png"))); // NOI18N
        JbtnImportarLinks.setText("Importar Excel");
        JbtnImportarLinks.setBorder(javax.swing.BorderFactory.createCompoundBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255), 3), javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED)));
        JbtnImportarLinks.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        JbtnImportarLinks.setFocusPainted(false);
        JbtnImportarLinks.setPressedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Importar Pressed.png"))); // NOI18N
        JbtnImportarLinks.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/importar roll.png"))); // NOI18N
        JbtnImportarLinks.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                JbtnImportarLinksMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                JbtnImportarLinksMouseExited(evt);
            }
        });
        JbtnImportarLinks.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                JbtnImportarLinksActionPerformed(evt);
            }
        });

        jBtnNormalizadorLinks.setBackground(new java.awt.Color(247, 247, 247));
        jBtnNormalizadorLinks.setFont(new java.awt.Font("Segoe UI", 3, 18)); // NOI18N
        jBtnNormalizadorLinks.setForeground(new java.awt.Color(0, 0, 0));
        jBtnNormalizadorLinks.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Normalizar.png"))); // NOI18N
        jBtnNormalizadorLinks.setText("Normalizador");
        jBtnNormalizadorLinks.setBorder(javax.swing.BorderFactory.createCompoundBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255), 3), javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED)));
        jBtnNormalizadorLinks.setBorderPainted(false);
        jBtnNormalizadorLinks.setContentAreaFilled(false);
        jBtnNormalizadorLinks.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jBtnNormalizadorLinks.setEnabled(false);
        jBtnNormalizadorLinks.setFocusPainted(false);
        jBtnNormalizadorLinks.setPressedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Normalizar Pressed.png"))); // NOI18N
        jBtnNormalizadorLinks.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/normalizar Roll.png"))); // NOI18N
        jBtnNormalizadorLinks.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jBtnNormalizadorLinksMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jBtnNormalizadorLinksMouseExited(evt);
            }
        });
        jBtnNormalizadorLinks.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBtnNormalizadorLinksActionPerformed(evt);
            }
        });

        jBtnExportarLinks.setBackground(new java.awt.Color(247, 247, 247));
        jBtnExportarLinks.setFont(new java.awt.Font("Segoe UI", 3, 18)); // NOI18N
        jBtnExportarLinks.setForeground(new java.awt.Color(0, 0, 0));
        jBtnExportarLinks.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Exportar.png"))); // NOI18N
        jBtnExportarLinks.setText("Exportar Excel");
        jBtnExportarLinks.setBorder(javax.swing.BorderFactory.createCompoundBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255), 3), javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED)));
        jBtnExportarLinks.setBorderPainted(false);
        jBtnExportarLinks.setContentAreaFilled(false);
        jBtnExportarLinks.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jBtnExportarLinks.setEnabled(false);
        jBtnExportarLinks.setFocusPainted(false);
        jBtnExportarLinks.setPressedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Exportar Pressed.png"))); // NOI18N
        jBtnExportarLinks.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Exportar Roll.png"))); // NOI18N
        jBtnExportarLinks.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jBtnExportarLinksMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jBtnExportarLinksMouseExited(evt);
            }
        });
        jBtnExportarLinks.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBtnExportarLinksActionPerformed(evt);
            }
        });

        jScrollPane3.setOpaque(false);

        tablaLinks.setFont(tablaLinks.getFont());
        tablaLinks.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {

            }
        ));
        tablaLinks.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_ALL_COLUMNS);
        tablaLinks.setAutoscrolls(false);
        tablaLinks.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        tablaLinks.setOpaque(false);
        jScrollPane3.setViewportView(tablaLinks);

        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/UltraFibra.png"))); // NOI18N

        javax.swing.GroupLayout jPanelLinkPagosLayout = new javax.swing.GroupLayout(jPanelLinkPagos);
        jPanelLinkPagos.setLayout(jPanelLinkPagosLayout);
        jPanelLinkPagosLayout.setHorizontalGroup(
            jPanelLinkPagosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelLinkPagosLayout.createSequentialGroup()
                .addGroup(jPanelLinkPagosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanelLinkPagosLayout.createSequentialGroup()
                        .addGap(69, 69, 69)
                        .addComponent(JbtnImportarLinks, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(88, 88, 88)
                        .addComponent(jBtnNormalizadorLinks, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(85, 85, 85)
                        .addComponent(jBtnExportarLinks, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanelLinkPagosLayout.createSequentialGroup()
                        .addGap(250, 250, 250)
                        .addComponent(jLabel2)))
                .addContainerGap(117, Short.MAX_VALUE))
            .addGroup(jPanelLinkPagosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanelLinkPagosLayout.createSequentialGroup()
                    .addGap(18, 18, 18)
                    .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 1040, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addContainerGap(12, Short.MAX_VALUE)))
        );
        jPanelLinkPagosLayout.setVerticalGroup(
            jPanelLinkPagosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelLinkPagosLayout.createSequentialGroup()
                .addGap(21, 21, 21)
                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 128, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addGroup(jPanelLinkPagosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(JbtnImportarLinks, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jBtnNormalizadorLinks, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jBtnExportarLinks, javax.swing.GroupLayout.PREFERRED_SIZE, 83, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(589, Short.MAX_VALUE))
            .addGroup(jPanelLinkPagosLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanelLinkPagosLayout.createSequentialGroup()
                    .addGap(298, 298, 298)
                    .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 366, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addContainerGap(178, Short.MAX_VALUE)))
        );

        jTabbedPane1.addTab("PAGOS POR LINKS", jPanelLinkPagos);

        JbtnImportarPMC.setBackground(new java.awt.Color(247, 247, 247));
        JbtnImportarPMC.setFont(new java.awt.Font("Segoe UI", 3, 18)); // NOI18N
        JbtnImportarPMC.setForeground(new java.awt.Color(0, 0, 0));
        JbtnImportarPMC.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/importar.png"))); // NOI18N
        JbtnImportarPMC.setText("Importar TXT");
        JbtnImportarPMC.setBorder(javax.swing.BorderFactory.createCompoundBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255), 3), javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED)));
        JbtnImportarPMC.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        JbtnImportarPMC.setFocusPainted(false);
        JbtnImportarPMC.setPressedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Importar Pressed.png"))); // NOI18N
        JbtnImportarPMC.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/importar roll.png"))); // NOI18N
        JbtnImportarPMC.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                JbtnImportarPMCMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                JbtnImportarPMCMouseExited(evt);
            }
        });
        JbtnImportarPMC.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                JbtnImportarPMCActionPerformed(evt);
            }
        });

        jBtnNormalizadorPMC.setBackground(new java.awt.Color(247, 247, 247));
        jBtnNormalizadorPMC.setFont(new java.awt.Font("Segoe UI", 3, 18)); // NOI18N
        jBtnNormalizadorPMC.setForeground(new java.awt.Color(0, 0, 0));
        jBtnNormalizadorPMC.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Normalizar.png"))); // NOI18N
        jBtnNormalizadorPMC.setText("Normalizador");
        jBtnNormalizadorPMC.setBorder(javax.swing.BorderFactory.createCompoundBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255), 3), javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED)));
        jBtnNormalizadorPMC.setBorderPainted(false);
        jBtnNormalizadorPMC.setContentAreaFilled(false);
        jBtnNormalizadorPMC.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jBtnNormalizadorPMC.setEnabled(false);
        jBtnNormalizadorPMC.setFocusPainted(false);
        jBtnNormalizadorPMC.setPressedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Normalizar Pressed.png"))); // NOI18N
        jBtnNormalizadorPMC.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/normalizar Roll.png"))); // NOI18N
        jBtnNormalizadorPMC.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jBtnNormalizadorPMCMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jBtnNormalizadorPMCMouseExited(evt);
            }
        });
        jBtnNormalizadorPMC.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBtnNormalizadorPMCActionPerformed(evt);
            }
        });

        jBtnExportarPMC.setBackground(new java.awt.Color(247, 247, 247));
        jBtnExportarPMC.setFont(new java.awt.Font("Segoe UI", 3, 18)); // NOI18N
        jBtnExportarPMC.setForeground(new java.awt.Color(0, 0, 0));
        jBtnExportarPMC.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Exportar.png"))); // NOI18N
        jBtnExportarPMC.setText("Exportar Excel");
        jBtnExportarPMC.setBorder(javax.swing.BorderFactory.createCompoundBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 255, 255), 3), javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED)));
        jBtnExportarPMC.setBorderPainted(false);
        jBtnExportarPMC.setContentAreaFilled(false);
        jBtnExportarPMC.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jBtnExportarPMC.setEnabled(false);
        jBtnExportarPMC.setFocusPainted(false);
        jBtnExportarPMC.setPressedIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Exportar Pressed.png"))); // NOI18N
        jBtnExportarPMC.setRolloverIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/Exportar Roll.png"))); // NOI18N
        jBtnExportarPMC.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jBtnExportarPMCMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jBtnExportarPMCMouseExited(evt);
            }
        });
        jBtnExportarPMC.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jBtnExportarPMCActionPerformed(evt);
            }
        });

        jScrollPane4.setOpaque(false);

        tablaPMC.setFont(tablaPMC.getFont());
        tablaPMC.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {

            }
        ));
        tablaPMC.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_ALL_COLUMNS);
        tablaPMC.setAutoscrolls(false);
        tablaPMC.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        tablaPMC.setOpaque(false);
        jScrollPane4.setViewportView(tablaPMC);

        jLabel3.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Imagenes/UltraFibra.png"))); // NOI18N

        javax.swing.GroupLayout jPanelLinkPagos1Layout = new javax.swing.GroupLayout(jPanelLinkPagos1);
        jPanelLinkPagos1.setLayout(jPanelLinkPagos1Layout);
        jPanelLinkPagos1Layout.setHorizontalGroup(
            jPanelLinkPagos1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelLinkPagos1Layout.createSequentialGroup()
                .addGroup(jPanelLinkPagos1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanelLinkPagos1Layout.createSequentialGroup()
                        .addGap(69, 69, 69)
                        .addComponent(JbtnImportarPMC, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(88, 88, 88)
                        .addComponent(jBtnNormalizadorPMC, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(85, 85, 85)
                        .addComponent(jBtnExportarPMC, javax.swing.GroupLayout.PREFERRED_SIZE, 237, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanelLinkPagos1Layout.createSequentialGroup()
                        .addGap(250, 250, 250)
                        .addComponent(jLabel3)))
                .addContainerGap(117, Short.MAX_VALUE))
            .addGroup(jPanelLinkPagos1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanelLinkPagos1Layout.createSequentialGroup()
                    .addGap(18, 18, 18)
                    .addComponent(jScrollPane4, javax.swing.GroupLayout.PREFERRED_SIZE, 1040, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addContainerGap(12, Short.MAX_VALUE)))
        );
        jPanelLinkPagos1Layout.setVerticalGroup(
            jPanelLinkPagos1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanelLinkPagos1Layout.createSequentialGroup()
                .addGap(21, 21, 21)
                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 128, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addGroup(jPanelLinkPagos1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(JbtnImportarPMC, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jBtnNormalizadorPMC, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jBtnExportarPMC, javax.swing.GroupLayout.PREFERRED_SIZE, 83, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(589, Short.MAX_VALUE))
            .addGroup(jPanelLinkPagos1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanelLinkPagos1Layout.createSequentialGroup()
                    .addGap(298, 298, 298)
                    .addComponent(jScrollPane4, javax.swing.GroupLayout.PREFERRED_SIZE, 366, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addContainerGap(178, Short.MAX_VALUE)))
        );

        jTabbedPane1.addTab("PAGO MIS CUENTAS", jPanelLinkPagos1);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jTabbedPane1, javax.swing.GroupLayout.Alignment.TRAILING)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addComponent(jTabbedPane1)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void JbtnImportarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_JbtnImportarActionPerformed
        importar(jTableDatosExcel2);
        jBtnNormalizador.setEnabled(true);
        jBtnNormalizador.setBorderPainted(true);
        jBtnNormalizador.setContentAreaFilled(true);
    }//GEN-LAST:event_JbtnImportarActionPerformed

    private void jBtnNormalizadorActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBtnNormalizadorActionPerformed
        int flag = 1;

        if (rBtnAbiertos.isSelected()) {
            flag = 1;
        } else if (rBtnCerrados.isSelected()) {
            flag = 2;
        }

        if (jdc_fecha.getDate() == null) {
            JOptionPane.showMessageDialog(null, "Debe Ingresar una fecha de vencimiento");

        } else {
            int resp = JOptionPane.showConfirmDialog(rootPane, "Selecciono Correctamente el tipo de Debitos?", "Confirmacion", YES_NO_OPTION);

            if (resp == 0) {
                Date date = jdc_fecha.getDate();
                LocalDate fecha = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();

                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyLLdd");
                String strFecha = fecha.format(formatter);

                FormatoDebitosAutomaticos nom = new FormatoDebitosAutomaticos(flag, strFecha);
                nom.normalizar(jTableDatosExcel2);
                jBtnNormalizador.setEnabled(false);
                jBtnNormalizador.setBorderPainted(false);
                jBtnNormalizador.setContentAreaFilled(false);
                jBtnExportar.setEnabled(true);
                jBtnExportar.setBorderPainted(true);
                jBtnExportar.setContentAreaFilled(true);

            }

        }

    }//GEN-LAST:event_jBtnNormalizadorActionPerformed

    private void jBtnExportarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBtnExportarActionPerformed
        exportar(jTableDatosExcel2);
        jdc_fecha.setCalendar(null);
        jTableDatosExcel2.setModel(new DefaultTableModel());
        jBtnExportar.setEnabled(false);
        jBtnExportar.setBorderPainted(false);
        jBtnExportar.setContentAreaFilled(false);
    }//GEN-LAST:event_jBtnExportarActionPerformed

    private void JbtnImportarMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_JbtnImportarMouseEntered
        JbtnImportar.setFont(new Font("Segoe UI", 3, 20));
    }//GEN-LAST:event_JbtnImportarMouseEntered

    private void JbtnImportarMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_JbtnImportarMouseExited
        JbtnImportar.setFont(new Font("Segoe UI", 1, 18));
    }//GEN-LAST:event_JbtnImportarMouseExited

    private void jBtnNormalizadorMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnNormalizadorMouseEntered
        if (jBtnNormalizador.isEnabled()) {
            jBtnNormalizador.setFont(new Font("Segoe UI", 3, 20));
        }

    }//GEN-LAST:event_jBtnNormalizadorMouseEntered

    private void jBtnNormalizadorMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnNormalizadorMouseExited
        if (jBtnNormalizador.isEnabled()) {
            jBtnNormalizador.setFont(new Font("Segoe UI", 1, 18));
        }

    }//GEN-LAST:event_jBtnNormalizadorMouseExited

    private void jBtnExportarMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnExportarMouseEntered
        if (jBtnExportar.isEnabled()) {
            jBtnExportar.setFont(new Font("Segoe UI", 3, 20));
        }

    }//GEN-LAST:event_jBtnExportarMouseEntered

    private void jBtnExportarMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnExportarMouseExited
        if (jBtnExportar.isEnabled()) {
            jBtnExportar.setFont(new Font("Segoe UI", 1, 18));
        }
    }//GEN-LAST:event_jBtnExportarMouseExited

    private void rBtnAbiertosMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_rBtnAbiertosMouseEntered
        if (!rBtnAbiertos.isSelected()) {
            rBtnAbiertos.setFont(new Font("Segoe UI", 3, 16));
        }

    }//GEN-LAST:event_rBtnAbiertosMouseEntered

    private void rBtnAbiertosMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_rBtnAbiertosMouseExited
        rBtnAbiertos.setForeground(new Color(0, 0, 0));
        rBtnAbiertos.setFont(new Font("Segoe UI", 1, 14));
    }//GEN-LAST:event_rBtnAbiertosMouseExited

    private void rBtnCerradosMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_rBtnCerradosMouseEntered
        if (!rBtnCerrados.isSelected()) {
            rBtnCerrados.setFont(new Font("Segoe UI", 3, 16));
        }

    }//GEN-LAST:event_rBtnCerradosMouseEntered

    private void rBtnCerradosMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_rBtnCerradosMouseExited
        rBtnCerrados.setForeground(new Color(0, 0, 0));
        rBtnCerrados.setFont(new Font("Segoe UI", 1, 14));
    }//GEN-LAST:event_rBtnCerradosMouseExited

    private void rBtnCerradosActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_rBtnCerradosActionPerformed

        rBtnCerrados.setBorderPainted(true);
        rBtnAbiertos.setBorderPainted(false);
        rBtnCerrados.setFont(new Font("Segoe UI", 1, 14));

    }//GEN-LAST:event_rBtnCerradosActionPerformed

    private void rBtnAbiertosActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_rBtnAbiertosActionPerformed
        rBtnAbiertos.setBorderPainted(true);
        rBtnCerrados.setBorderPainted(false);
        rBtnAbiertos.setFont(new Font("Segoe UI", 1, 14));

    }//GEN-LAST:event_rBtnAbiertosActionPerformed

    private void JbtnImportarLinksMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_JbtnImportarLinksMouseEntered
        JbtnImportarLinks.setFont(new Font("Segoe UI", 3, 20));
    }//GEN-LAST:event_JbtnImportarLinksMouseEntered

    private void JbtnImportarLinksMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_JbtnImportarLinksMouseExited
        JbtnImportarLinks.setFont(new Font("Segoe UI", 1, 18));
    }//GEN-LAST:event_JbtnImportarLinksMouseExited

    private void JbtnImportarLinksActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_JbtnImportarLinksActionPerformed
        importar(tablaLinks);
        jBtnNormalizadorLinks.setEnabled(true);
        jBtnNormalizadorLinks.setBorderPainted(true);
        jBtnNormalizadorLinks.setContentAreaFilled(true);
    }//GEN-LAST:event_JbtnImportarLinksActionPerformed

    private void jBtnNormalizadorLinksMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnNormalizadorLinksMouseEntered
        if (jBtnNormalizadorLinks.isEnabled()) {
            jBtnNormalizadorLinks.setFont(new Font("Segoe UI", 3, 20));
        }
    }//GEN-LAST:event_jBtnNormalizadorLinksMouseEntered

    private void jBtnNormalizadorLinksMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnNormalizadorLinksMouseExited
        if (jBtnNormalizadorLinks.isEnabled()) {
            jBtnNormalizadorLinks.setFont(new Font("Segoe UI", 1, 18));
        }
    }//GEN-LAST:event_jBtnNormalizadorLinksMouseExited

    private void jBtnNormalizadorLinksActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBtnNormalizadorLinksActionPerformed
        FormatoMacroClicks nom = new FormatoMacroClicks();
        nom.normalizar(tablaLinks);

        jBtnNormalizadorLinks.setEnabled(false);
        jBtnNormalizadorLinks.setBorderPainted(false);
        jBtnNormalizadorLinks.setContentAreaFilled(false);
        jBtnExportarLinks.setEnabled(true);
        jBtnExportarLinks.setBorderPainted(true);
        jBtnExportarLinks.setContentAreaFilled(true);

    }//GEN-LAST:event_jBtnNormalizadorLinksActionPerformed

    private void jBtnExportarLinksMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnExportarLinksMouseEntered
        if (jBtnExportarLinks.isEnabled()) {
            jBtnExportarLinks.setFont(new Font("Segoe UI", 3, 20));
        }
    }//GEN-LAST:event_jBtnExportarLinksMouseEntered

    private void jBtnExportarLinksMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnExportarLinksMouseExited
        if (jBtnExportarLinks.isEnabled()) {
            jBtnExportarLinks.setFont(new Font("Segoe UI", 1, 18));
        }
    }//GEN-LAST:event_jBtnExportarLinksMouseExited

    private void jBtnExportarLinksActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBtnExportarLinksActionPerformed
        exportar(tablaLinks);
        tablaLinks.setModel(new DefaultTableModel());
        jBtnExportarLinks.setEnabled(false);
        jBtnExportarLinks.setBorderPainted(false);
        jBtnExportarLinks.setContentAreaFilled(false);
    }//GEN-LAST:event_jBtnExportarLinksActionPerformed

    private void JbtnImportarPMCMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_JbtnImportarPMCMouseEntered
        JbtnImportarPMC.setFont(new Font("Segoe UI", 3, 20));
    }//GEN-LAST:event_JbtnImportarPMCMouseEntered

    private void JbtnImportarPMCMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_JbtnImportarPMCMouseExited
        JbtnImportarPMC.setFont(new Font("Segoe UI", 1, 18));
    }//GEN-LAST:event_JbtnImportarPMCMouseExited

    private void JbtnImportarPMCActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_JbtnImportarPMCActionPerformed
        try {
            importarTxt(tablaPMC);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Home.class.getName()).log(Level.SEVERE, null, ex);
        }
        jBtnNormalizadorPMC.setEnabled(true);
        jBtnNormalizadorPMC.setBorderPainted(true);
        jBtnNormalizadorPMC.setContentAreaFilled(true);
    }//GEN-LAST:event_JbtnImportarPMCActionPerformed

    private void jBtnNormalizadorPMCMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnNormalizadorPMCMouseEntered
         if (jBtnNormalizadorPMC.isEnabled()) {
            jBtnNormalizadorPMC.setFont(new Font("Segoe UI", 3, 20));
        }
    }//GEN-LAST:event_jBtnNormalizadorPMCMouseEntered

    private void jBtnNormalizadorPMCMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnNormalizadorPMCMouseExited
        if (jBtnNormalizadorPMC.isEnabled()) {
            jBtnNormalizadorPMC.setFont(new Font("Segoe UI", 1, 18));
        }
    }//GEN-LAST:event_jBtnNormalizadorPMCMouseExited

    private void jBtnNormalizadorPMCActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBtnNormalizadorPMCActionPerformed
        FormatoHomebanking nom = new FormatoHomebanking();
        nom.normalizar(tablaPMC);

        jBtnNormalizadorPMC.setEnabled(false);
        jBtnNormalizadorPMC.setBorderPainted(false);
        jBtnNormalizadorPMC.setContentAreaFilled(false);
        jBtnExportarPMC.setEnabled(true);
        jBtnExportarPMC.setBorderPainted(true);
        jBtnExportarPMC.setContentAreaFilled(true);
    }//GEN-LAST:event_jBtnNormalizadorPMCActionPerformed

    private void jBtnExportarPMCMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnExportarPMCMouseEntered
         if (jBtnExportarPMC.isEnabled()) {
            jBtnExportarPMC.setFont(new Font("Segoe UI", 3, 20));
        }
    }//GEN-LAST:event_jBtnExportarPMCMouseEntered

    private void jBtnExportarPMCMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jBtnExportarPMCMouseExited
          if (jBtnExportarPMC.isEnabled()) {
            jBtnExportarPMC.setFont(new Font("Segoe UI", 1, 18));
        }
    }//GEN-LAST:event_jBtnExportarPMCMouseExited

    private void jBtnExportarPMCActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jBtnExportarPMCActionPerformed
        exportar(tablaPMC);
        tablaPMC.setModel(new DefaultTableModel());
        jBtnExportarPMC.setEnabled(false);
        jBtnExportarPMC.setBorderPainted(false);
        jBtnExportarPMC.setContentAreaFilled(false);
    }//GEN-LAST:event_jBtnExportarPMCActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {

        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Home().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton JbtnImportar;
    private javax.swing.JButton JbtnImportarLinks;
    private javax.swing.JButton JbtnImportarPMC;
    private javax.swing.ButtonGroup buttonGroup1;
    private javax.swing.JButton jBtnExportar;
    private javax.swing.JButton jBtnExportarLinks;
    private javax.swing.JButton jBtnExportarPMC;
    private javax.swing.JButton jBtnNormalizador;
    private javax.swing.JButton jBtnNormalizadorLinks;
    private javax.swing.JButton jBtnNormalizadorPMC;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabelLogo;
    private javax.swing.JPanel jPaneDebitos;
    private javax.swing.JPanel jPanelLinkPagos;
    private javax.swing.JPanel jPanelLinkPagos1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JScrollPane jScrollPane4;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JTable jTableDatosExcel2;
    public static com.toedter.calendar.JDateChooser jdc_fecha;
    private javax.swing.JRadioButton rBtnAbiertos;
    private javax.swing.JRadioButton rBtnCerrados;
    private javax.swing.JTable tablaLinks;
    private javax.swing.JTable tablaPMC;
    // End of variables declaration//GEN-END:variables

    class FondoPanel extends JPanel {

        private Image imagen;

        @Override
        public void paint(Graphics g) {
            imagen = new ImageIcon(getClass().getResource("/Imagenes/Fondo.jpg")).getImage();
            g.drawImage(imagen, 0, 0, getWidth(), getHeight(), this);
            setOpaque(false);
            super.paint(g);
        }
    }

}
