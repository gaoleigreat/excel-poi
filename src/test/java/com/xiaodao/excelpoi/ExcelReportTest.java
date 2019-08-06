package com.xiaodao.excelpoi;


import com.alibaba.fastjson.JSONObject;
import com.xiaodao.excelpoi.element.EObject;
import com.xiaodao.excelpoi.model.Coordinate;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class ExcelReportTest extends ExcelPoiApplicationTests{

    @Test
    public void getReplaceParam() {
        try {
            ExcelReport excelReport = new ExcelReport("G:/code/excel-poi/src/main/resources/template.xlsx");
          List<Coordinate> coordinates = excelReport.getReplaceParam();
            for (Coordinate coordinate:coordinates) {
               System.out.println(JSONObject.toJSONString(coordinate));
            }
        } catch (IOException e) {
            log.error("init excel faill",e.getMessage());
        }

    }


    @Test
    public void tetsReplaceParame() throws IOException {
        ExcelReport excelReport = new ExcelReport("C:/Users/xiaodao//Desktop/xiaodao/code/git/excel-poi/src/main/resources/template.xlsx");
        List<EObject> myEobjects = new ArrayList<>();
        for(int i =0;i<100;i++){
            MyEobject myEobject = new MyEobject();
            myEobject.setProjectName("工程名"+i);
            myEobject.setName("高磊"+i);
            myEobject.setAddress("陕西"+i);
            myEobject.setAge("27"+i);
            myEobjects.add(myEobject);
        }
       excelReport.insertIntoTable(0,11,myEobjects,true);
        excelReport.insertPicture();
        excelReport.generate("C:/Users/xiaodao//Desktop/xiaodao/code/git/excel-poi/src/main/resources/templateOut.xlsx");
    }
}