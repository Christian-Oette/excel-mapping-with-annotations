# Excel Mapping with annotated Pojo

```java
@RequiredArgsConstructor
@Builder
public class TestPojo {
    @ExcelColumn(name = "Welcome Text")
    private final String text;

    @ExcelColumn(name = "Test Number", dataFormat="#,#0.0")
    private final BigDecimal decimalNumber;

    @ExcelColumn(name = "Created at", dataFormat = "m/d/yy")
    private final LocalDate localDate = LocalDate.now();
}
```

## How to use

```java
    var workbookBuilder = new ExcelWorkbookBuilder(ExcelFileFormat.XLSX);
    List<Object> data = ...
    Workbook wb = workbookBuilder
    .addSheet("sheet name", data)
    .build();

    try  (OutputStream fileOut = new FileOutputStream("output/output.xlsx")) {
        wb.write(fileOut);
    }
```