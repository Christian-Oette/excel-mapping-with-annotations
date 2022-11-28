package com.christianoette.mapper;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public enum ExcelFileFormat {
    XLS{
        @Override
        Workbook createNewWorkbook() {
            return new HSSFWorkbook();
        }
    },
    XLSX {
        @Override
        Workbook createNewWorkbook() {
            return new XSSFWorkbook();
        }
    }
    ;

    abstract Workbook createNewWorkbook();
}
