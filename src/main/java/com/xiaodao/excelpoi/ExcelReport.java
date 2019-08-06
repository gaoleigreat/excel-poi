package com.xiaodao.excelpoi;

import com.xiaodao.excelpoi.element.EObject;
import com.xiaodao.excelpoi.model.Coordinate;
import com.xiaodao.excelpoi.util.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
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

    public void replaceParame(List<Coordinate> coordinates,EObject eObject){
        for (Coordinate coordinate:coordinates) {
            String cellValue = eObject.getTextByKey(coordinate.getValue().toString());
            xsf.getSheetAt(coordinate.getSheetNumber()).getRow(coordinate.getRowNumber()).getCell(coordinate.getLineNumber()).setCellValue(cellValue);
        }
    }

    public void insertIntoTable(int sheetNumber,int rowNumber,List<EObject> eObjects){
        List<Coordinate> coordinates = new ArrayList<>();
        XSSFRow row = xsf.getSheetAt(sheetNumber).getRow(rowNumber);

        xsf.getSheetAt(sheetNumber).shiftRows(rowNumber+1, xsf.getSheetAt(sheetNumber).getLastRowNum(),eObjects.size(), true, false);

        for (int i=0;i<eObjects.size();i++) {
            xsf.getSheetAt(sheetNumber).getRow(rowNumber+1+i);
        }


        int totalLineNumber = row.getPhysicalNumberOfCells();
        for (int lineNumber = 0; lineNumber < totalLineNumber; lineNumber++) {
            XSSFCell cell = row.getCell(lineNumber);
            String cellValue = cell.getStringCellValue();
            Matcher matcher = ExcelUtil.matcher(cellValue);
            while (matcher.find()) {
                coordinates.add(new Coordinate(sheetNumber, rowNumber, lineNumber, matcher.group(1)));
            }
        }
        for (int i=0;i<eObjects.size();i++) {
            for (Coordinate coordinate:coordinates ) {
                xsf.getSheetAt(coordinate.getSheetNumber()).createRow(coordinate.getRowNumber()+i).createCell(coordinate.getLineNumber()).setCellValue(eObjects.get(i).getTextByKey(coordinate.getValue().toString()));
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
        xsf.getSheetAt(0).shiftRows(12,20,9);
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

}
