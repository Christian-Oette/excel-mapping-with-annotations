package com.christianoette.mapper;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.List;
import java.util.Optional;

import static org.assertj.core.api.Assertions.assertThat;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

class ExcelWorkbookBuilderTest {

    @Test
    @Disabled // run manually only
    public void manuallyCreateOutputFile() {
        ExcelWorkbookBuilder workbookBuilder = new ExcelWorkbookBuilder(ExcelFileFormat.XLSX);
        List<Object> data = List.of(
            new TestPojo("Hello World", null, null),
            new TestPojo("Hello Excel", 12, new BigDecimal("27.8999"))
        );
        Workbook wb = workbookBuilder
            .addSheet("one", data)
            .addSheet("two", data)
            .build();
        writeToFile(wb);
    }

    @SneakyThrows
    private void writeToFile(final Workbook wb) {
        try (OutputStream fileOut = new FileOutputStream("output/output.xlsx")) {
            wb.write(fileOut);
        }
    }

    @Test
    public void thatWorkbookBuilderWorksCorrectly() {
        // given
        var workbookBuilder = new ExcelWorkbookBuilder(ExcelFileFormat.XLSX);
        var testPojo = new TestPojo("Hello Excel", 12, new BigDecimal("27.8999"));
        var testDate = LocalDate.of(2022, 1, 2);
        var testDateTime = LocalDateTime.of(testDate, LocalTime.NOON);
        testPojo.setLocalDate(testDate);
        testPojo.setLocalDateTime(testDateTime);
        var data = List.of(
            testPojo
        );

        // when
        var wb = workbookBuilder
            .addSheet("sheet", data)
            .build();

        // then
        assertThat(wb).isNotNull();
        assertThat(wb.getNumberOfSheets()).isEqualTo(1);
        var sheet = wb.getSheetAt(0);
        assertThat(sheet.getSheetName()).isEqualTo("sheet");
        assertThat(strValue(sheet, 0, 0)).isEqualTo("Welcome Text");
        assertThat(strValue(sheet, 0, 1)).isEqualTo("Nr1");
        assertThat(strValue(sheet, 0, 2)).isEqualTo("Nr2");
        assertThat(strValue(sheet, 0, 3)).isEqualTo("Created at");
        assertThat(strValue(sheet, 0, 4)).isEqualTo("Created at with time");
        assertThat(strValue(sheet, 1, 0)).isEqualTo("Hello Excel");
        assertThat(strValue(sheet, 1, 1)).isEqualTo("12");
        assertThat(numericValue(sheet, 1, 2)).isBetween(27.0,29.0);
        assertThat(numericValue(sheet, 1, 3)).isBetween(44000.0,45000.0);
        assertThat(numericValue(sheet, 1, 4)).isBetween(44000.0,45000.0);
    }

    private String strValue(final Sheet sheet, final int row, final int column) {
        return Optional.ofNullable(sheet)
            .map(s -> s.getRow(row))
            .map(r -> r.getCell(column))
            .map(Cell::getStringCellValue)
            .orElse(null);
    }

    private Double numericValue(final Sheet sheet, final int row, final int column) {
        return Optional.ofNullable(sheet)
            .map(s -> s.getRow(row))
            .map(r -> r.getCell(column))
            .map(Cell::getNumericCellValue)
            .orElse(null);
    }
}