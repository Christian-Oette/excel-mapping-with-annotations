package com.christianoette.mapper;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.NumberFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

import com.christianoette.mapper.annotation.ExcelColumn;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@RequiredArgsConstructor
public class ExcelWorkbookBuilder {
    private final ExcelFileFormat format;
    private final List<ExcelSheetData> sheets = new ArrayList<>();

    public ExcelWorkbookBuilder addSheet(String sheetName, List<Object> data) {
        sheets.add(new ExcelSheetData(sheetName, data));
        return this;
    }

    public Workbook build() {
        ExcelFileFormat format = Objects.requireNonNull(this.format, "ExcelFormat must not be null");
        Workbook workbook = format.createNewWorkbook();
        createSheets(workbook);
        return workbook;
    }

    private void createSheets(final Workbook workbook) {
        for (ExcelSheetData excelSheetData : sheets) {
            Sheet sheet = workbook.createSheet(excelSheetData.getName());
            createTitleRow(sheet, excelSheetData);
            createRowsInSheet(sheet, excelSheetData);
            autoSizeColumns(sheet, excelSheetData.getData().size());
        }
    }

    private void autoSizeColumns(final Sheet sheet, final int size) {
        for (int i = 0; i < size; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    @SneakyThrows
    private void createTitleRow(final Sheet sheet, final ExcelSheetData excelSheetData) {
        Row row = sheet.createRow(0);
        int columnIndex = 0;
        Optional<Object> firstItemOptional = excelSheetData.getData().stream().findFirst();
        if (firstItemOptional.isPresent()) {
            Object item = firstItemOptional.get();
            Field[] allFields = item.getClass().getDeclaredFields();
            for (Field field : allFields) {
                ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                if (annotation != null) {
                    Cell cell = row.createCell(columnIndex);
                    columnIndex++;
                    writeCellValue(annotation.name(), cell, annotation);
                }
            }
        }
    }

    private void createRowsInSheet(final Sheet sheet, final ExcelSheetData excelSheetData) {
        int rowNumber = 1;
        for (Object item : excelSheetData.getData()) {
            fillRow(sheet, item, rowNumber);
            rowNumber++;
        }
    }

    @SneakyThrows
    private void fillRow(final Sheet sheet, final Object columnData, int rowNumber) {
        Row row = sheet.createRow(rowNumber);
        int columnIndex = 0;
        Field[] allFields = columnData.getClass().getDeclaredFields();
        for (Field field : allFields) {
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation != null) {
                field.setAccessible(true);
                Object fieldValue = field.get(columnData);
                Cell cell = row.createCell(columnIndex);
                columnIndex++;
                writeCellValue(fieldValue, cell, annotation);
            }
        }
    }

    private void writeCellValue(final Object fieldValue, final Cell cell, final ExcelColumn annotation) {
        String defaultFormat = null;
        if (fieldValue instanceof Double) {
            cell.setCellValue((double) fieldValue);
        } else if (fieldValue instanceof BigDecimal) {
            cell.setCellValue(((BigDecimal) fieldValue).doubleValue());
        } else if (fieldValue instanceof Boolean) {
            cell.setCellValue((boolean) fieldValue);
        } else if (fieldValue instanceof LocalDate) {
            cell.setCellValue((LocalDate) fieldValue);
        } else if (fieldValue instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) fieldValue);
        } else {
            cell.setCellValue(String.valueOf(fieldValue));
        }
        applyDataFormatIfSet(cell, annotation.dataFormat(), defaultFormat);
    }

    private void applyDataFormatIfSet(final Cell cell, String dataFormat, String defaultFormat) {
        if (dataFormat != null && dataFormat.length() > 0) {
            applyDataFormat(cell, dataFormat);
        } else if (defaultFormat!=null && defaultFormat.length() > 0){
            applyDataFormat(cell, defaultFormat);
        }
    }

    private void applyDataFormat(final Cell cell, final String dataFormat) {
        Workbook wb = cell.getSheet().getWorkbook();
        DataFormat format = wb.createDataFormat();
        CellStyle style = wb.createCellStyle();
        style.setDataFormat(format.getFormat(dataFormat));
        cell.setCellStyle(style);
    }

    protected String getLocalizedBigDecimalValue(BigDecimal input) {
        final NumberFormat numberFormat = NumberFormat.getNumberInstance(Locale.US);
        numberFormat.setGroupingUsed(true);
        numberFormat.setMaximumFractionDigits(2);
        numberFormat.setMinimumFractionDigits(2);
        return numberFormat.format(input);
    }
}
