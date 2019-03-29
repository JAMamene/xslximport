package testo.xlsx.importtest;

import com.opencsv.CSVWriter;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

public class SheetToCSV implements XSSFSheetXMLHandler.SheetContentsHandler {
    private boolean firstCellOfRow;
    private int currentRow = -1;
    private int currentCol = -1;
    private int minColumns = -1;
    private List<String> line;
    private CSVWriter csvWriter;

    public SheetToCSV(CSVWriter csvFile) {
        line = new ArrayList<>();
        csvWriter = csvFile;
    }

    static XMLReader fetchSheetParser(SharedStringsTable sst, CSVWriter csvFile) throws SAXException, ParserConfigurationException {
        XMLReader sheetParser = SAXHelper.newXMLReader();
        ContentHandler handler = new XSSFSheetXMLHandler(null, null, sst, new SheetToCSV(csvFile), new DataFormatter(), false);
        sheetParser.setContentHandler(handler);
        return sheetParser;
    }

    private void outputMissingRows(int number) {
        for (int i = 0; i < number; i++) {
            for (int j = 0; j < minColumns; j++) {
                line.add("");
            }
            csvWriter.writeNext(line.toArray(new String[0]),false);
            line = new ArrayList<>();
        }
    }

    @Override
    public void startRow(int rowNum) {
        outputMissingRows(rowNum - currentRow - 1);
        firstCellOfRow = true;
        currentRow = rowNum;
        currentCol = -1;
    }

    @Override
    public void endRow(int rowNum) {
        for (int i = currentCol; i < minColumns; i++) {
            line.add("");
        }
        csvWriter.writeNext(line.toArray(new String[0]),false);
        line = new ArrayList<>();
    }

    @Override
    public void cell(String cellReference, String formattedValue,
                     XSSFComment comment) {
        if (firstCellOfRow) {
            firstCellOfRow = false;
        } else {
            line.add("");
        }

        if (cellReference == null) {
            cellReference = new CellAddress(currentRow, currentCol).formatAsString();
        }

        int thisCol = (new CellReference(cellReference)).getCol();
        int missedCols = thisCol - currentCol - 1;
        for (int i = 0; i < missedCols; i++) {
            line.add("");
        }
        currentCol = thisCol;

        line.add(formattedValue);
    }
}
