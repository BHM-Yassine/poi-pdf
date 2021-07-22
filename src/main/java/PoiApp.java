import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.*;

public class PoiApp {

    public static void main(String [] args) throws IOException, InvalidFormatException {
        FileInputStream fis = new FileInputStream("test.docx");
        FileOutputStream out = new FileOutputStream("test2.docx");

        XWPFDocument doc = new XWPFDocument(OPCPackage.open(fis));

        XWPFTable table = doc.getTables().get(0);

        XWPFTableRow row = table.getRow(1); // First row
        row.getCell(0).setText("AA");
        row.getCell(1).setText("BB");
        row.getCell(2).setText("CC");

        XWPFTableRow row1 = table.getRow(3); // First row
        row1.getCell(0).setText("AA");
        row1.getCell(1).setText("BB");
        row1.getCell(2).setText("CC");

        XWPFTableRow row2 = table.getRow(5); // First row
        row2.getCell(0).setText("AA");
        row2.getCell(1).setText("BB");
        row2.getCell(2).setText("CC");

        XWPFTableRow row3 = table.getRow(7); // First row
        row3.getCell(0).setText("Comment");

        doc.write(out);
    }
}
