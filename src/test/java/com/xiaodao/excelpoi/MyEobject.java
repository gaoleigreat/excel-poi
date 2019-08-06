package com.xiaodao.excelpoi;

import com.xiaodao.excelpoi.element.EObject;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class MyEobject extends EObject {

    private String projectName;

    private String age;

    private String name ;

    private String address;



    @Override
    public Object getValByKey(String key) {
        if ("projectName".equals(key)){
            return this.getProjectName();
        }
        if ("age".equals(key)){
            return this.getAge();
        }
        if ("name".equals(key)){
            return this.getName();
        }
        if ("address".equals(key)){
            return this.getAddress();
        }
        return null;
    }
}
