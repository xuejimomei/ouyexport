package com.example.oyexport.excel;

import com.example.oyexport.common.ExcelConstants;
import com.example.oyexport.common.ExcelExportUtil;
import com.example.oyexport.common.ExportContext;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.dom4j.Element;

import java.util.List;
import java.util.Map;

public enum EExportType {
    FIXED(ExcelConstants.TABLE_TYPE_1) {
        @Override
        public ExportContext action(ExportContext context, List<Map<String, Object>> dataList, List<Map<String, Object>> dataListext, Integer headRowEnd, Element structureElement, SXSSFSheet sheet) {
            Integer dataUsedRow = context.getDataUsedRow();
            Map<String, Integer> indexList = context.getIndexList();
            // 无动态列 所有表头按照配置写死
            ExcelExportUtil.createFixedHead(context, 0, headRowEnd, sheet, structureElement);
            for (int i = 0; i < dataList.size(); i++) {
                context.setDataUsedRow(++dataUsedRow);//每条数据另起一行
                SXSSFRow row1 = sheet.createRow(dataUsedRow);
                System.out.println("+++++++++++++++++++++++++++++++++第" + i + "条数据");
                Map<String, Object> map = dataList.get(i);
                //行列 都存在     直接填写数据
                for (Map.Entry<String, Object> entry : map.entrySet()) {
                    Integer col = indexList.get(entry.getKey());
                    if (col != null) {// 有些字段并未出现在报表中 加此判断 否则空指针
                        SXSSFCell cell = row1.createCell(col);
                        ExcelExportUtil.cellFillData(context, entry.getKey(), entry.getValue(), cell);
                    }
                }
            }
            return context;
        }
    }, FLEX(ExcelConstants.TABLE_TYPE_2) {
        @Override
        public ExportContext action(ExportContext context, List<Map<String, Object>> dataList, List<Map<String, Object>> dataListext, Integer headRowEnd, Element structureElement, SXSSFSheet sheet) {
            String firstHead = context.getFirstHead();
            String flexHead = context.getFlexHead();
            Integer dataUsedRow = context.getDataUsedRow();
            Integer readCount = context.getReadCount();
            Map<String, Integer> indexList = context.getIndexList();

            // 有动态列 的情况
            for (int i = 0; i < dataList.size(); i++) {
                System.out.println("+++++++++++++++++++++++++++++++++第" + i + "条数据");
                Map<String, Object> map = dataList.get(i);
                String rowKey = firstHead + "_" + map.get(firstHead); //行
                String colKey = flexHead + "_" + map.get(flexHead); //root 动态列
                if (indexList.get(colKey) == null) {
                    //列不存在
                    context = ExcelExportUtil.createHeadAndFillData(context, 0, headRowEnd, sheet, structureElement, map);
                } else if (indexList.get(rowKey) != null && indexList.get(colKey) == null) {
                    //列存在 行不存在   新建一行写入数据
                    int row = dataUsedRow + 1;
                    for (Map.Entry<String, Object> entry : map.entrySet()) {
                        String key = flexHead + "_" + map.get(flexHead) + "_" + entry.getKey();
                        Integer col = indexList.get(key);
                        SXSSFCell cell = sheet.createRow(row).createCell(col);
                        ExcelExportUtil.cellFillData(context, entry.getKey(), entry.getValue(), cell);
                    }
                } else if (indexList.get(rowKey) != null && indexList.get(colKey) != null) {
                    //行列 都存在     直接填写数据
                    Integer row = indexList.get(rowKey);
                    for (Map.Entry<String, Object> entry : map.entrySet()) {
                        String key = flexHead + "_" + map.get(flexHead) + "_" + entry.getKey(); //行
                        Integer col = indexList.get(key);
                        if (col != null) {// 有些字段并未出现在报表中 加此判断 否则空指针
                            SXSSFCell cell = sheet.getRow(row).createCell(col);
                            ExcelExportUtil.cellFillData(context, entry.getKey(), entry.getValue(), cell);
                        }
                    }
                }
                context.setReadCount(++readCount);
            }
            //2. 额外字段填入数据
            if (dataListext.size() > 0) {
                // TODO: 2019/11/11 mysql 用join 解决吧,使用拼接的不好分页
                ExcelExportUtil.setExtData(context, dataListext, sheet);
            }
            return context;
        }
    };
    private Integer type;

    EExportType(Integer t) {
        this.type = t;
    }

    public static EExportType getType(Integer t) {
        for (EExportType c : EExportType.values()) {
            if (c.getType().equals(t)) {
                return c;
            }
        }
        return null;
    }

    public Integer getType() {
        return type;
    }

    public abstract ExportContext action(ExportContext context, List<Map<String, Object>> dataList, List<Map<String, Object>> dataListext, Integer headRowEnd, Element structureElement, SXSSFSheet sheet);
}
