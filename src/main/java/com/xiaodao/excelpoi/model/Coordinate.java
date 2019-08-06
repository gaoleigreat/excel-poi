package com.xiaodao.excelpoi.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Coordinate {
    /**
     * sheet坐标
     */
    private Integer sheetNumber;

    /**
     * 行坐标
     */
    private Integer rowNumber;

    /**
     * 列坐标
     */
    private Integer lineNumber;

    /**
     * 值
     */
    private Object value;
}
