package com.lego.excel.poi;

import com.lego.excel.poi.element.EPic;
import com.lego.excel.poi.element.EObject;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class MyEobject extends EObject {

    private String projectName;

    private Integer age;

    private String name;

    private String address;

    private EPic picture;


    @Override
    public Object getValByKey(String key) {
        if ("projectName".equals(key)) {
            return this.getProjectName();
        }
        if ("age".equals(key)) {
            return this.getAge();
        }
        if ("name".equals(key)) {
            return this.getName();
        }
        if ("address".equals(key)) {
            return this.getAddress();
        }
        if ("picture".equals(key)) {
            return this.getPicture();
        }
        return null;
    }
}
