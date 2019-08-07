package com.lego.excel.poi.element;

public abstract class EObject {

    public abstract Object getValByKey(String key);

    public boolean isPic(String key) {
        return this.getValByKey(key) instanceof EPic;
    }

    public String getTextByKey(String key) {
        if (!isPic(key)) {
            return (String) this.getValByKey(key);
        } else {
            return null;
        }
    }

    public EPic getPicByKey(String key) {
        if (isPic(key)) {
            return (EPic) this.getValByKey(key);
        } else {
            return null;
        }
    }
}
