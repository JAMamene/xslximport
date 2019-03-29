package testo.xlsx.importtest;

import com.monitorjbl.xlsx.StreamingReader;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.util.Iterator;

import static testo.xlsx.importtest.PrintContentHandler.fetchSheetParser;


public enum ImportType implements Importable {

    FASTEXCEL {
        @Override
        public void execute(File file, PrintStream out) throws IOException {
            InputStream is = new FileInputStream(file);
            ReadableWorkbook wb = new ReadableWorkbook(is);
            wb.getSheets().forEach(s ->
                    {
                        try {
                            s.openStream().forEach(r ->
                                    r.forEach(c -> {
                                        if (c != null) {
                                            out.println(c.asString());
                                        }
                                    })
                            );
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
            );
        }
    },

    EXCEL_STREAMING {
        @Override
        public void execute(File file, PrintStream out) throws IOException {
            InputStream is = new FileInputStream(file);
            Workbook workbook = StreamingReader.builder()
                    .rowCacheSize(10)
                    .bufferSize(2048)
                    .open(is);

            for (Sheet sheet : workbook) {
                for (Row r : sheet) {
                    for (Cell c : r) {
                        out.println(c.getStringCellValue());
                    }
                }

            }
        }
    },

    POI_EVENTDRIVEN {
        @Override
        public void execute(File file, PrintStream out) throws IOException {
            try {
                OPCPackage pkg = OPCPackage.open(file.getAbsolutePath());
                XSSFReader r = new XSSFReader(pkg);
                SharedStringsTable sst = r.getSharedStringsTable();
                XMLReader parser = fetchSheetParser(sst, out);
                readSheets(r, parser);
            } catch (OpenXML4JException | ParserConfigurationException | SAXException e) {
                e.printStackTrace();
            }
        }
    },

    POI {
        @Override
        public void execute(File file, PrintStream out) throws IOException {
            Workbook wb = WorkbookFactory.create(file);
            DataFormatter dataFormatter = new DataFormatter();
            wb.forEach(sheet ->
                    sheet.forEach(row ->
                            row.forEach(cell ->
                                    out.println(dataFormatter.formatCellValue(cell))
                            )
                    )
            );
        }
    },

    CSV {
        @Override
        public File beforeExecute(File file) {
            try {
                File csvFile = File.createTempFile("tmp", ".csv", null);
                CSVWriter csvWriter = new CSVWriter(new FileWriter(csvFile));
                OPCPackage opcPackage = OPCPackage.open(file.getAbsolutePath());
                XSSFReader xssfReader = new XSSFReader(opcPackage);
                SharedStringsTable sharedStringsTable = xssfReader.getSharedStringsTable();
                XMLReader parser = SheetToCSV.fetchSheetParser(sharedStringsTable, csvWriter);
                readSheets(xssfReader, parser);
                return csvFile;
            } catch (SAXException | OpenXML4JException | IOException | ParserConfigurationException e) {
                e.printStackTrace();
            }
            return null;
        }

        @Override
        public void execute(File file, PrintStream out) throws IOException {
            Reader reader = new BufferedReader(new FileReader(file));
            CSVReader csvReader = new CSVReader(reader);
            String[] nextRecord;
            while ((nextRecord = csvReader.readNext()) != null) {
                for (String cell : nextRecord) {
                    out.println(cell);
                }
            }
        }
    };

    private static void readSheets(XSSFReader xssfReader, XMLReader parser) throws IOException, InvalidFormatException, SAXException {
        Iterator<InputStream> sheets = xssfReader.getSheetsData();
        while (sheets.hasNext()) {
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }
    }

    public static ImportType getEnumByString(String code) {
        for (ImportType importType : ImportType.values()) {
            if (code.equals(importType.name())) return importType;
        }
        return null;
    }
}


