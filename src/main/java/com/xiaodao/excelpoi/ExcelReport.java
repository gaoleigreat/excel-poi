package com.xiaodao.excelpoi;

import com.xiaodao.excelpoi.element.EObject;
import com.xiaodao.excelpoi.model.Coordinate;
import com.xiaodao.excelpoi.util.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;

@Slf4j
public class ExcelReport {
    private String templatePath;
    private XSSFWorkbook xsf = null;
    private OutputStream os = null;
    private InputStream is = null;

    /**
     * 构造方法
     *
     * @param templatePath
     * @throws IOException
     */
    public ExcelReport(String templatePath) throws IOException {
        this.templatePath = templatePath;
        init();
    }

    /**
     * 初始化XSSFWorkbook
     *
     * @throws IOException
     */
    public void init() {
        try {
            is = ExcelUtil.getFileStream(this.templatePath);
            xsf = new XSSFWorkbook(is);
        } catch (Exception e) {
            log.error("excel 表格初始化失败", e.getMessage());
        }
    }

    /**
     * 获取Excel中所有可以替换的变量的值和坐标
     *
     * @return List<Coordinate>
     */
    public List<Coordinate> getReplaceParam() {
        List<Coordinate> coordinates = new ArrayList<>();
        int totalSheetNumber = xsf.getNumberOfSheets();
        for (int sheetNumber = 0; sheetNumber < totalSheetNumber; sheetNumber++) {
            XSSFSheet sheet = xsf.getSheetAt(sheetNumber);
            int totalRowNumber = sheet.getPhysicalNumberOfRows();
            for (int rowNumber = 0; rowNumber < totalRowNumber; rowNumber++) {
                XSSFRow row = sheet.getRow(rowNumber);
                int totalLineNumber = row.getPhysicalNumberOfCells();
                for (int lineNumber = 0; lineNumber < totalLineNumber; lineNumber++) {
                    XSSFCell cell = row.getCell(lineNumber);
                    String cellValue = cell.getStringCellValue();
                    Matcher matcher = ExcelUtil.matcher(cellValue);
                    while (matcher.find()) {
                        coordinates.add(new Coordinate(sheetNumber, rowNumber, lineNumber, matcher.group(1)));
                    }

                }

            }

        }
        return coordinates;
    }


    /**
     * 对Excel中的值进行替换
     */

    public void replaceParame(List<Coordinate> coordinates, EObject eObject) {
        for (Coordinate coordinate : coordinates) {
            String cellValue = eObject.getTextByKey(coordinate.getValue().toString());
            xsf.getSheetAt(coordinate.getSheetNumber()).getRow(coordinate.getRowNumber()).getCell(coordinate.getLineNumber()).setCellValue(cellValue);
        }
    }

    public void insertIntoTable(int sheetNumber, int rowNumber, List<EObject> eObjects, boolean hasBorder) {
        List<Coordinate> coordinates = new ArrayList<>();
        XSSFRow row = xsf.getSheetAt(sheetNumber).getRow(rowNumber);
        int totalLineNumber = row.getPhysicalNumberOfCells();
        xsf.getSheetAt(sheetNumber).shiftRows(rowNumber + 1, xsf.getSheetAt(sheetNumber).getLastRowNum(), eObjects.size() - 1, true, false);
        CellStyle style = xsf.createCellStyle();
        style.setBorderBottom(CellStyle.BORDER_THIN); // 下边框
        style.setBorderLeft(CellStyle.BORDER_THIN);// 左边框
        style.setBorderTop(CellStyle.BORDER_THIN);// 上边框
        style.setBorderRight(CellStyle.BORDER_THIN);// 右边框
        for (int i = 0; i < eObjects.size() - 1; i++) {
            xsf.getSheetAt(sheetNumber).createRow(rowNumber + i + 1);
            if (hasBorder) {
                for (int j = 0; j < totalLineNumber; j++) {
                    xsf.getSheetAt(sheetNumber).getRow(rowNumber + i + 1).createCell(j);
                    xsf.getSheetAt(sheetNumber).getRow(rowNumber + i + 1).getCell(j).setCellStyle(style);
                }
            }
        }


        for (int lineNumber = 0; lineNumber < totalLineNumber; lineNumber++) {
            XSSFCell cell = row.getCell(lineNumber);
            String cellValue = cell.getStringCellValue();
            Matcher matcher = ExcelUtil.matcher(cellValue);
            while (matcher.find()) {
                coordinates.add(new Coordinate(sheetNumber, rowNumber, lineNumber, matcher.group(1)));
            }
        }

        for (int i = 0; i < eObjects.size(); i++) {
            for (Coordinate coordinate : coordinates) {
                xsf.getSheetAt(coordinate.getSheetNumber()).getRow(coordinate.getRowNumber() + i).getCell(coordinate.getLineNumber()).setCellValue(eObjects.get(i).getTextByKey(coordinate.getValue().toString()));
            }
        }

    }

    public boolean generate(String outDocPath) throws IOException {
        os = new FileOutputStream(outDocPath);
        xsf.write(os);
        ExcelUtil.close(os);
        ExcelUtil.close(is);
        return true;
    }


    public boolean generate1(String outDocPath) throws IOException {
        xsf.getSheetAt(0).shiftRows(12, 20, 9);
        xsf.getSheetAt(0).createRow(12);
        xsf.getSheetAt(0).createRow(13);
        xsf.getSheetAt(0).createRow(14);
        xsf.getSheetAt(0).createRow(15);
        xsf.getSheetAt(0).createRow(16);
        xsf.getSheetAt(0).createRow(17);
        xsf.getSheetAt(0).createRow(18);
        xsf.getSheetAt(0).createRow(19);
        os = new FileOutputStream(outDocPath);
        xsf.write(os);
        ExcelUtil.close(os);
        ExcelUtil.close(is);
        return true;
    }


    public void setCellValue(int sheetNumber, int rowNumber, int lineNumber, Object value) {
        if (value instanceof Integer || value instanceof Double) {
            xsf.getSheetAt(sheetNumber).getRow(rowNumber).getCell(lineNumber).setCellValue(Double.valueOf(value.toString()));
        } else {
            xsf.getSheetAt(sheetNumber).getRow(rowNumber).getCell(lineNumber).setCellValue(value.toString());
        }

    }

    public void insertPicture() throws IOException {
        XSSFDrawing xssfDrawing = xsf.getSheetAt(0).createDrawingPatriarch();
        xsf.addPicture(new FileInputStream("C:/Users/xiaodao/Desktop/1.png"), Workbook.PICTURE_TYPE_PNG);
        ClientAnchor anchor = new XSSFClientAnchor(1, 1, 1, 1,
                8, 11, 9, 12);
        xssfDrawing.createPicture(anchor,0);
    }

}
