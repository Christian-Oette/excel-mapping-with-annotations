package com.christianoette.mapper;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

import com.christianoette.mapper.annotation.ExcelColumn;
import lombok.Builder;
import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
@Builder
public class TestPojo {
    @ExcelColumn(name = "Welcome Text")
    private final String text;

    @ExcelColumn(name = "Nr1")
    private final Integer aNumber;

    @ExcelColumn(name = "Nr2", dataFormat="#,#0.0")
    private final BigDecimal decimalNumber;

    @ExcelColumn(name = "Created at", dataFormat = "m/d/yy")
    private final LocalDate localDate = LocalDate.now();

    @ExcelColumn(name = "Created at with time", dataFormat = "m/d/yy h:mm")
    private final LocalDateTime localDateTime = LocalDateTime.now();
}
