package testo.xlsx.importtest;

import java.io.PrintStream;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.binary.XSSFBSheetHandler.SheetContentsHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

public class PrintContentHandler implements SheetContentsHandler {
    private boolean firstCellOfRow;
    private int currentRow = -1;
    private int currentCol = -1;
    private int minColumns;
    private PrintStream out;

    public PrintContentHandler(PrintStream out, int minColumns) {
        this.out = out;
        this.minColumns = minColumns;
    }

    static XMLReader fetchSheetParser(SharedStringsTable sst, PrintStream out) throws SAXException, ParserConfigurationException {
        XMLReader sheetParser = SAXHelper.newXMLReader();
        ContentHandler handler = new XSSFSheetXMLHandler(null, null, sst, new PrintContentHandler(out, 10), new DataFormatter(), false);
        sheetParser.setContentHandler(handler);
        return sheetParser;
    }

    private void outputMissingRows(int number) {
        for(int i = 0; i < number; ++i) {
            for(int j = 0; j < this.minColumns; ++j) {
                System.out.print(',');
            }

            System.out.println();
        }

    }

    public void startRow(int rowNum) {
        this.outputMissingRows(rowNum - this.currentRow - 1);
        this.firstCellOfRow = true;
        this.currentRow = rowNum;
        this.currentCol = -1;
    }

    public void endRow(int rowNum) {
        for(int i = this.currentCol; i < this.minColumns; ++i) {
            this.out.print(',');
        }

        this.out.print('\n');
    }

    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        if (this.firstCellOfRow) {
            this.firstCellOfRow = false;
        } else {
            this.out.print(',');
        }

        this.out.print(" ");

        try {
            Double.parseDouble(formattedValue);
            this.out.print(formattedValue);
        } catch (NumberFormatException var5) {
            this.out.print(formattedValue);
        }

    }

    public void hyperlinkCell(String s, String s1, String s2, String s3, XSSFComment xssfComment) {
    }
}
