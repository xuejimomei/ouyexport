package com.example.oyexport.common;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class IfCondition {
    String ifSymbol;      //符号  > >=  <  <=  =
    String ifTarget;      //比较对象    50    50%
    XSSFCellStyle cellStyle; //单元格样式

    public String getIfSymbol() {
        return ifSymbol;
    }

    public void setIfSymbol(String ifSymbol) {
        this.ifSymbol = ifSymbol;
    }

    public String getIfTarget() {
        return ifTarget;
    }

    public void setIfTarget(String ifTarget) {
        this.ifTarget = ifTarget;
    }

    public XSSFCellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(XSSFCellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}
