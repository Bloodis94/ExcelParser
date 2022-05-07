package com.github.excelparser.parser.iface;

import com.github.excelparser.Iface.Destination;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

@NoArgsConstructor
public abstract class ExcelParserAbstract {

    public <T extends Destination> ArrayList<T> destinationParser(Workbook origin, Class<T> destination) {

        ArrayList<T> destinations = new ArrayList<>();

        Sheet sheet = origin.getSheetAt(0);
        int rowCount = sheet.getPhysicalNumberOfRows();
        ArrayList<String> columnNames = getColumnNames(sheet);

        for (int i = 1; i <= rowCount - 1; i++) {

            Row row = sheet.getRow(i);
            T destinationObj = null;

            try {
                destinationObj = destination.getConstructor().newInstance();
            } catch (InstantiationException e) {
                throw new RuntimeException(e);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            } catch (InvocationTargetException e) {
                throw new RuntimeException(e);
            } catch (NoSuchMethodException e) {
                throw new RuntimeException(e);
            }

            Field[] fields = destinationObj.getClass().getDeclaredFields();

            for (Field field : fields) {

                try {
                    field.set(destinationObj, getColumnValue(columnNames.indexOf(field.getName()), row));
                } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
            }

            destinations.add(destinationObj);
        }

        return destinations;
    }

    public Workbook originParser(Class<Workbook> origin, ArrayList<Destination> destinations) {

        Workbook workbook = null;

        try {
            workbook = origin.getConstructor().newInstance();
        } catch (InstantiationException e) {
            throw new RuntimeException(e);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        } catch (InvocationTargetException e) {
            throw new RuntimeException(e);
        } catch (NoSuchMethodException e) {
            throw new RuntimeException(e);
        }

        for (Destination destination : destinations) {

            Row row = workbook.getSheetAt(0).getRow(0);
            ArrayList<Field> fields = new ArrayList<>(List.of(destination.getClass().getDeclaredFields()));

            for (Field field : fields) {

                row.createCell(fields.indexOf(field)).setCellValue(field.getName());
            }

            Row destinationRow = workbook.getSheetAt(0).createRow(destinations.indexOf(destination));

            for (Field field : fields) {

                try {
                    destinationRow.createCell(fields.indexOf(field)).setCellValue(String.valueOf(field.get(destination)));
                } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
            }
        }

        return workbook;
    }

    private ArrayList<String> getColumnNames(Sheet sheet) {

        ArrayList<String> columnNames = new ArrayList<>();

        Row row = sheet.getRow(0);
        int cellCount = row.getPhysicalNumberOfCells();

        for (int i = 0; i <= cellCount - 1; i++) {

            Cell cell = row.getCell(i);

            columnNames.add(cell.getStringCellValue());
        }

        return columnNames;
    }

    private String getColumnValue(int columnIndex, Row row) {

        return row.getCell(columnIndex).getStringCellValue();
    }
}
