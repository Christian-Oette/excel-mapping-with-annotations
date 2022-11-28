package com.christianoette.mapper;

import java.util.List;
import java.util.Objects;

import lombok.Getter;
import org.apache.poi.ss.util.WorkbookUtil;

public class ExcelSheetData {
    @Getter
    private final String name;
    @Getter
    private final List<Object> data;

    ExcelSheetData(final String name, final List<Object> dataDtos) {
        this.name = WorkbookUtil.createSafeSheetName(Objects.requireNonNull(name, "Sheet name must not be null"));
        this.data = dataDtos;
    }
}
