package com.exceljdbc.table;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class TableManager {

    private final Workbook workbook;
    private final String filePath;

    public TableManager(Workbook workbook, String filePath) {
        this.workbook = workbook;
        this.filePath = filePath;
    }

    public List<Map<String, Object>> getTableData(String tableName) {
        Sheet sheet = workbook.getSheet(tableName);
        if (sheet == null) return null;

        List<Map<String, Object>> rows = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.iterator();
        if (!rowIterator.hasNext()) return rows;

        Row headerRow = rowIterator.next();
        List<String> columns = new ArrayList<>();
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            columns.add(cell != null ? cell.getStringCellValue().trim() : "");
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Map<String, Object> rowMap = new LinkedHashMap<>();
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                rowMap.put(columns.get(i), getCellValue(cell));
            }
            rows.add(rowMap);
        }
        return rows;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue();
                double num = cell.getNumericCellValue();
                if (num == Math.floor(num) && !Double.isInfinite(num) && num <= Long.MAX_VALUE && num >= Long.MIN_VALUE)
                    return (long) num;
                return num;
            case BOOLEAN: return cell.getBooleanCellValue();
            case FORMULA:
                try {
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    CellValue cv = evaluator.evaluate(cell);
                    switch (cv.getCellType()) {
                        case NUMERIC: return cv.getNumberValue();
                        case STRING: return cv.getStringValue();
                        case BOOLEAN: return cv.getBooleanValue();
                        default: return null;
                    }
                } catch (Exception e) {
                    try { return cell.getStringCellValue(); } catch (Exception ex) { return null; }
                }
            default: return null;
        }
    }

    public void insertRow(String tableName, Map<String, Object> values) {
        Sheet sheet = workbook.getSheet(tableName);
        if (sheet == null) return;
        List<String> columns = getHeaderColumns(sheet);
        int newRowIdx = sheet.getLastRowNum() + 1;
        Row newRow = sheet.createRow(newRowIdx);
        for (int i = 0; i < columns.size(); i++) {
            Cell cell = newRow.createCell(i);
            Object val = values.get(columns.get(i));
            if (val != null) setCellValue(cell, val);
        }
    }

    public void updateRows(String tableName, List<Map<String, Object>> rows) {
        Sheet sheet = workbook.getSheet(tableName);
        if (sheet == null) return;
        List<String> columns = getHeaderColumns(sheet);
        for (int i = sheet.getLastRowNum(); i > 0; i--) {
            Row row = sheet.getRow(i);
            if (row != null) sheet.removeRow(row);
        }
        int rowIdx = 1;
        for (Map<String, Object> rowMap : rows) {
            Row row = sheet.createRow(rowIdx++);
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = row.createCell(i);
                Object val = rowMap.get(columns.get(i));
                if (val != null) setCellValue(cell, val);
            }
        }
    }

    private void setCellValue(Cell cell, Object value) {
        if (value == null) { cell.setBlank(); return; }
        if (value instanceof String) cell.setCellValue((String) value);
        else if (value instanceof Long) cell.setCellValue((Long) value);
        else if (value instanceof Integer) cell.setCellValue((Integer) value);
        else if (value instanceof Double) cell.setCellValue((Double) value);
        else if (value instanceof Float) cell.setCellValue((Float) value);
        else if (value instanceof Date) cell.setCellValue((Date) value);
        else if (value instanceof Boolean) cell.setCellValue((Boolean) value);
        else cell.setCellValue(value.toString());
    }

    private List<String> getHeaderColumns(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return Collections.emptyList();
        List<String> columns = new ArrayList<>();
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            columns.add(cell != null ? cell.getStringCellValue().trim() : "");
        }
        return columns;
    }

    public void createTable(String tableName, List<String> columnDefs) {
        Sheet sheet = workbook.createSheet(tableName);
        Row header = sheet.createRow(0);
        for (int i = 0; i < columnDefs.size(); i++) {
            String colName = columnDefs.get(i).trim().split("\\s+")[0];
            header.createCell(i).setCellValue(colName);
        }
    }

    public void dropTable(String tableName) {
        int index = workbook.getSheetIndex(tableName);
        if (index >= 0) workbook.removeSheetAt(index);
    }

    public List<String> getTableNames() {
        List<String> names = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) names.add(workbook.getSheetName(i));
        return names;
    }

    public void save(String filePath) throws IOException {
        File file = new File(filePath);
        File parent = file.getParentFile();
        if (parent != null && !parent.exists()) parent.mkdirs();
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
            fos.flush();
        }
    }

    public void save() throws IOException {
        save(filePath);
    }
}