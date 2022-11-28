package com.christianoette.mapper;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.List;

import lombok.SneakyThrows;
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
        try  (OutputStream fileOut = new FileOutputStream("output/output.xlsx")) {
            wb.write(fileOut);
        }
    }

}