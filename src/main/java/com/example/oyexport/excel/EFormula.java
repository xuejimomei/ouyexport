package com.example.oyexport.excel;

import com.example.oyexport.common.ExcelConstants;
import com.example.oyexport.common.ExcelExportUtil;
import com.example.oyexport.common.ExportContext;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.dom4j.Element;

import java.util.Objects;

public enum EFormula {
    /**
     * 总计  平均的行数使用了dataUsedRow,数据填写完毕后dataUsedRow 不应该再改变
     * 所以添加了 dataUsedRowExtCount字段
     */
    TOTAL("total") {
        @Override
        public ExportContext setFormula(ExportContext context, Integer headRowEnd, SXSSFSheet sheet, Element node) {
            Integer dataUsedRow = context.getDataUsedRow();
            Integer headUsedCol = context.getHeadUsedCol();
            Integer dataUsedRowExtCount = context.getDataUsedRowExount();
            //设置 "总计" 表头(纵向)
            SXSSFRow row = sheet.createRow(dataUsedRow + dataUsedRowExtCount + 1);//dataUsedRow不在变化  含增加通过dataUsedRowExtCount 实现
            SXSSFCell headCell = row.createCell(0);
            headCell.setCellValue(node.attributeValue(ExcelConstants.HEAD_NAME));
            ExcelExportUtil.setHeadStyleOfFormula(context, node, headCell);
            for (int i = 1; i < headUsedCol; i++) { //从第二列开始求 列求和
                SXSSFCell dataCell = row.createCell(i);//设置公式前，一定要先建立表格
                //设置样式  总计  平均值
                ExcelExportUtil.setDataStyleOfFormula(context, node, dataCell);
                String colString = CellReference.convertNumToColString(i);  //长度转成ABC列
                int i1 = headRowEnd + 2;//SUM(B1:B5)    1 和5 表示的是execel 看到的第一行  第五行   与其他逻辑不同(行列从0开始)
                int i2 = dataUsedRow + 1;//dataUsedRow - dataUsedRowExtCount 不统计其他统计求和所在的行
                String sumstring = "SUM(" + colString + i1 + ":" + colString + i2 + ")";//求和公式
                System.out.println(sumstring);
                dataCell.setCellFormula(sumstring);
            }
            context.setDataUsedRowExount(++dataUsedRowExtCount);
            return context;
        }
    }, AVERAGE("average") {
        @Override
        public ExportContext setFormula(ExportContext context, Integer headRowEnd, SXSSFSheet sheet, Element node) {
            Integer dataUsedRow = context.getDataUsedRow();
            Integer headUsedCol = context.getHeadUsedCol();
            Integer dataUsedRowExtCount = context.getDataUsedRowExount();
            //设置 "平均值" 表头(纵向)
            SXSSFRow row = sheet.createRow(dataUsedRow + dataUsedRowExtCount + 1);//dataUsedRow不在变化  含增加通过dataUsedRowExtCount 实现
            SXSSFCell headCell = row.createCell(0);
            headCell.setCellValue(node.attributeValue(ExcelConstants.HEAD_NAME));
            ExcelExportUtil.setHeadStyleOfFormula(context, node, headCell);
            for (int i = 1; i < headUsedCol; i++) { //从第二列开始求 列求和
                SXSSFCell dataCell = row.createCell(i);//设置公式前，一定要先建立表格
                ExcelExportUtil.setDataStyleOfFormula(context, node, dataCell);
                String colString = CellReference.convertNumToColString(i);  //长度转成ABC列
                int i1 = headRowEnd + 2;//SUM(B1:B5)    1 和5 表示的是execel 看到的第一行  第五行   与其他逻辑不同(行列从0开始)
                int i2 = dataUsedRow + 1;//dataUsedRow - dataUsedRowExtCount 不统计其他统计求和所在的行
                String range = "(" + colString + i1 + ":" + colString + i2 + ")";
                String sumstring = "SUM" + range + "/(COUNTBLANK" + range + "+" + "COUNTA" + range + ")";//AVERAGE 不统计空单元格 所用用sum 取平均值
                System.out.println(sumstring);
                dataCell.setCellFormula(sumstring);
            }
            context.setDataUsedRowExount(++dataUsedRowExtCount);
            return context;
        }
    };
    // TODO: 2019/10/31 后期的跟进要求
    private String type;

    EFormula(String t) {
        this.type = t;
    }

    public static EFormula getType(String t) {
        for (EFormula c : EFormula.values()) {
            if (Objects.equals(c.getType(), t)) {
                return c;
            }
        }
        return null;
    }

    public String getType() {
        return type;
    }

    public abstract ExportContext setFormula(ExportContext context, Integer headRowEnd, SXSSFSheet sheet, Element node);


}
