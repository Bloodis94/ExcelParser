package com.github.excelparser;

import com.github.excelparser.Iface.Destination;
import com.github.excelparser.parser.DefaultExcelParser;
import com.github.excelparser.parser.iface.ExcelParserAbstract;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class ExcelParserExecutor {

    ExcelParserAbstract parser;

    public ExcelParserExecutor(ExcelParserAbstract parser) {

        this.parser = parser;
    }

    public ExcelParserExecutor() {

        parser = new DefaultExcelParser();
    }

    public <T extends Destination> ArrayList<T> parseDestination(FileInputStream fis, Class<T> destination) {

        XSSFWorkbook origin = null;

        try {
            origin = new XSSFWorkbook(fis);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return parser.destinationParser(origin, destination);
    }

    public Workbook parseOrigin(Class<Workbook> origin, ArrayList<Destination> destinations) {

        return parser.originParser(origin, destinations);
    }
}
