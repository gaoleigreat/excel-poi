package com.xiaodao.excelpoi;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

@SpringBootApplication
public class ExcelPoiApplication {

    public static void main(String[] args) throws IOException {
        SpringApplication.run(ExcelPoiApplication.class, args);
        InputStream inputStream = new FileInputStream("G:/code/excel-poi/src/main/resources/template.xlsx");

        /**
         * 这里根据不同的excel类型
         * 可以选取不同的处理类：
         *          1.XSSFWorkbook
         *          2.HSSFWorkbook
         */
        // 获得工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        // 获得工作表
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rows = sheet.getPhysicalNumberOfRows();

        for (int i = 0; i < rows; i++) {

            // 获取第i行数据
            XSSFRow sheetRow = sheet.getRow(i);
            int o = sheetRow.getPhysicalNumberOfCells();
            System.out.println(o);
            for (int j = 0; j < o; j++) {
                XSSFCell cell = sheetRow.getCell(j);
                cell.getRow();
                // 调用toString方法获取内容
                System.out.println(cell);
            }



        }

    }
}