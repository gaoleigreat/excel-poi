package com.lego.excel.poi;


import com.alibaba.fastjson.JSONObject;
import com.lego.excel.poi.element.EObject;
import com.lego.excel.poi.element.EPic;
import com.lego.excel.poi.model.Coordinate;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class ExcelReportTest extends ExcelPoiApplicationTests {

    @Test
    public void getReplaceParam() {
        try {
            ExcelReport excelReport = new ExcelReport("G:/code/excel-poi/src/main/resources/template.xlsx");
            List<Coordinate> coordinates = excelReport.getReplaceParam();
            for (Coordinate coordinate : coordinates) {
                System.out.println(JSONObject.toJSONString(coordinate));
            }
        } catch (Exception e) {
            log.error("init excel faill", e.getMessage());
        }

    }


    @Test
    public void tetsReplaceParame() throws IOException {
        ExcelReport excelReport = new ExcelReport("G:/code/excel-poi/src/main/resources/template.xlsx");
        List<EObject> myEobjects = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            MyEobject myEobject = new MyEobject();
            myEobject.setProjectName("工程名" + i);
            myEobject.setName("高磊" + i);
            myEobject.setAddress("陕西" + i);
            myEobject.setAge(27 + i);
            EPic ePic = new EPic();
            ePic.setType("png");
            ePic.setUrl("http://114.115.233.31/group1/M00/00/00/wKgAX102zVeAAzg0AAAu-rWbsVk821_big.png");

            myEobject.setPicture(ePic);
            myEobjects.add(myEobject);
        }

        MyEobject myEobject = new MyEobject();
        myEobject.setProjectName("工程名");
        myEobject.setName("高磊");
        myEobject.setAddress("陕西");
        myEobject.setAge(27);
        EPic ePic = new EPic();
        ePic.setType("png");

        ePic.setUrl("http://114.115.233.31/group1/M00/00/00/wKgAX102zVeAAzg0AAAu-rWbsVk821_big.png");
        myEobject.setPicture(ePic);
        excelReport.insertIntoTable(0, 11, myEobjects, true);
        excelReport.replaceParame(myEobject);
        excelReport.generate("G:/code/excel-poi/src/main/resources/templateOut.xlsx");
    }
}