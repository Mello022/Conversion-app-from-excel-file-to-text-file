import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.util.Iterator;

public class Ayoub extends JFrame implements ActionListener{
    JButton button = new JButton("Importer");
    JButton button2 = new JButton("Exécuter");
    JComboBox combobox;
    String s;

    public static void main(String[] args) {
        new Ayoub().setVisible(true);
    }
    public Ayoub() {
        setSize(350, 300);
        setTitle("Convert From Excel To Text ");
        setLocationRelativeTo(null);
        setDefaultCloseOperation(EXIT_ON_CLOSE);

        button.addActionListener(this);
        button2.addActionListener(this);

        JLabel label = new JLabel("Importer le fichier excel :           ");
        JLabel label2 = new JLabel("Choisir le mode de conversion  :           ");

        JPanel p = new JPanel();
        p.add(label);
        p.add(button);
        p.setLayout(new FlowLayout(FlowLayout.LEFT));

        p.add(label2);
        String[] values = {"Host-object", "Groupe-object"} ;
        combobox = new JComboBox(values);

        p.add(combobox);
        combobox.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                System.out.println("Valeur: " + combobox.getSelectedItem().toString());
                s = combobox.getSelectedItem().toString();
            }
        });
        p.add(button2);
        setContentPane(p);

    }
    XSSFSheet sheet;
    FileInputStream file;
    JFileChooser fc;

    public void actionPerformed(ActionEvent e) {
        if(e.getSource() == button) {
            fc = new JFileChooser();
            if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION)
                try {
                    File selectedFile = fc.getSelectedFile();
                    System.out.println("Le chemin du fichier : " + selectedFile.getAbsolutePath());

                    file = new FileInputStream(new File(selectedFile.getAbsolutePath()));
                    XSSFWorkbook workbook = new XSSFWorkbook(file);
                    sheet = workbook.getSheetAt(0);

                } catch (Exception exception) {
                    exception.printStackTrace();
                }
        }
        if(e.getSource() == button2) {
            try {
                if(s == "Host-object"){
                    File f1 = new File("HOST.txt");
                    FileWriter fw2 = new FileWriter(f1);
                    Iterator<Row> rowIterator = sheet.iterator();
                    while (rowIterator.hasNext()) {
                        Row row = rowIterator.next();
                        Iterator<Cell> cellIterator = row.cellIterator();
                        Cell cell = cellIterator.next();

                        String cell1 = "object-network " + row.getCell(1).getStringCellValue();
                        cell.setCellValue(cell1);
                        fw2.write(cell.getStringCellValue());

                        String cell2 = "host " + row.getCell(2).getStringCellValue();
                        cell.setCellValue(cell2);
                        fw2.write("\n" + cell.getStringCellValue() + "\n");

                    }
                    file.close();
                    fw2.close();
                    System.out.println("Le fichier HOST.txt a été créer avec succès.");

                }
                else if (s == "Groupe-object"){
                    File f2 = new File("GROUPE.txt");
                    FileWriter fw = new FileWriter(f2);
                    fw.write("object-group network  IDEMIA"+ "\n");

                    Iterator<Row> rowIterator = sheet.iterator();
                    while (rowIterator.hasNext()) {
                        Row row = rowIterator.next();
                        Iterator<Cell> cellIterator = row.cellIterator();
                        Cell cell = cellIterator.next();

                        String cell3 = "network-object object " + row.getCell(1).getStringCellValue() + "\n";
                        cell.setCellValue(cell3);
                        fw.write(cell.getStringCellValue());

                    }
                    file.close();
                    fw.close();
                    System.out.println("Le fichier GROUPE.txt a été créer avec succès.");

                }

            } catch (Exception exception) {
                exception.printStackTrace();
            }
        }
    }
}
