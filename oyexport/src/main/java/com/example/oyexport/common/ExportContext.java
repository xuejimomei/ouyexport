package com.example.oyexport.common;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.dom4j.Element;

import java.util.HashMap;
import java.util.Map;

public class ExportContext {
    //xml 解析的通用变量
    private String tmplateCode;
    private Integer type;
    private Element structure;
    private Element structureExt;
    private String firstHead;
    private String flexHead;
    private Integer headRowCount;  //表头的行数
    //下面的是每张报表通用的变量
    private Integer headRowEnd;//表头使用列(下次新增表头 使用此列)
    private Integer headUsedCol = 0;//表头使用列(下次新增表头 使用此列)
    private Integer headUsedRow = -1;//表头使用行 已经使用到了此行
    private Integer dataUsedRow = null;//数据使用行 已经使用到了此行
    private Integer dataUsedRowExount = 0;
    private Integer readCount = 0;
    private SXSSFWorkbook workbook = new SXSSFWorkbook();
    private Map<String, Integer> indexList = new HashMap<>();// 记录 第一列的行数  其他列的列数 动态列父表头的起始列
    private Map<String, IfCondition> ifMap = new HashMap<>();//记录 ifCondition 的信息   key = 节点名称/字段名称    vlaue = if

    public ExportContext(String tmplateCode) {
        Element element = ExcelExportUtil.getElement(tmplateCode);
        this.structure = element.element(ExcelConstants.TABLE_STRUCTURE);
        this.structureExt = element.element(ExcelConstants.TABLE_STRUCTUREEXT);
        this.tmplateCode = tmplateCode;
        this.type = Integer.parseInt(element.element(ExcelConstants.TABLE_TYPE).getText());
        this.headRowCount = Integer.parseInt(element.element(ExcelConstants.TABLE_HEADROWCOUNT).getText());
        this.headRowEnd = headRowCount - 1;
        this.dataUsedRow = headRowCount - 1;
        this.firstHead = element.element(ExcelConstants.TABLE_FIRSTHEAD).getText();
        if (element.element(ExcelConstants.TABLE_FLEXHEAD) != null) {
            this.flexHead = element.element(ExcelConstants.TABLE_FLEXHEAD).getText();
        }
    }

    public String getTmplateCode() {
        return tmplateCode;
    }

    public void setTmplateCode(String tmplateCode) {
        this.tmplateCode = tmplateCode;
    }

    public Integer getType() {
        return type;
    }

    public void setType(Integer type) {
        this.type = type;
    }

    public Element getStructure() {
        return structure;
    }

    public void setStructure(Element structure) {
        this.structure = structure;
    }

    public Element getStructureExt() {
        return structureExt;
    }

    public void setStructureExt(Element structureExt) {
        this.structureExt = structureExt;
    }

    public String getFirstHead() {
        return firstHead;
    }

    public void setFirstHead(String firstHead) {
        this.firstHead = firstHead;
    }

    public String getFlexHead() {
        return flexHead;
    }

    public void setFlexHead(String flexHead) {
        this.flexHead = flexHead;
    }

    public Integer getHeadRowCount() {
        return headRowCount;
    }

    public void setHeadRowCount(Integer headRowCount) {
        this.headRowCount = headRowCount;
    }

    public Integer getHeadRowEnd() {
        return headRowEnd;
    }

    public void setHeadRowEnd(Integer headRowEnd) {
        this.headRowEnd = headRowEnd;
    }

    public Integer getHeadUsedCol() {
        return headUsedCol;
    }

    public void setHeadUsedCol(Integer headUsedCol) {
        this.headUsedCol = headUsedCol;
    }

    public Integer getHeadUsedRow() {
        return headUsedRow;
    }

    public void setHeadUsedRow(Integer headUsedRow) {
        this.headUsedRow = headUsedRow;
    }

    public Integer getDataUsedRow() {
        return dataUsedRow;
    }

    public void setDataUsedRow(Integer dataUsedRow) {
        this.dataUsedRow = dataUsedRow;
    }

    public Integer getDataUsedRowExount() {
        return dataUsedRowExount;
    }

    public void setDataUsedRowExount(Integer dataUsedRowExount) {
        this.dataUsedRowExount = dataUsedRowExount;
    }

    public Integer getReadCount() {
        return readCount;
    }

    public void setReadCount(Integer readCount) {
        this.readCount = readCount;
    }

    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(SXSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public Map<String, Integer> getIndexList() {
        return indexList;
    }

    public void setIndexList(Map<String, Integer> indexList) {
        this.indexList = indexList;
    }

    public Map<String, IfCondition> getIfMap() {
        return ifMap;
    }

    public void setIfMap(Map<String, IfCondition> ifMap) {
        this.ifMap = ifMap;
    }
}
