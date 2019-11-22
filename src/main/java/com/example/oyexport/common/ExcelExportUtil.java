package com.example.oyexport.common;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.util.ResourceUtils;

import java.math.BigDecimal;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Pattern;


/**
 * @Author: ouyangshaolei
 * @Date: 2019/7/4 14:56
 */
@Service
public class ExcelExportUtil {
    // 正则匹配所有 整数小数    (填入数据时数字不能以文字形式填入)
    private static final Pattern pattern = Pattern.compile("^(([^0][0-9]+|0)\\.([0-9]{1,2})$)|^(([^0][0-9]+|0)$)|^(([1-9]+)\\.([0-9]{1,2})$)|^(([1-9]+)$)");
    private static final Logger logger = LoggerFactory.getLogger(ExcelExportUtil.class);

    /**
     * 设置 公式求 和 平均值 的样式
     * 其他地方设置样式会导致 精度等样式失效
     *
     * @param context
     * @param node
     * @param dataCell
     * @return
     */
    public static void setDataStyleOfFormula(ExportContext context, Element node, SXSSFCell dataCell) {
        SXSSFWorkbook workbook = context.getWorkbook();
        String pr = node.attributeValue(ExcelConstants.HEAD_PRECISION); //小数精度
        String pe = node.attributeValue(ExcelConstants.HEAD_PERCENT); //百分数精度
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        //设置边框
        setAllBorderStyle(cellStyle);
        //设置 背景 字体颜色
        String fillColor = node.attributeValue(ExcelConstants.HEAD_DATAFILLCOLOR);//data  背景颜色
        String fontColor = node.attributeValue(ExcelConstants.HEAD_DATAFONTCOLOR);//data  字体颜色

        //背景景颜色
        if (fillColor != null) {
            cellStyle.setFillForegroundColor(fromStrToARGB(fillColor));
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        //设置字体 颜色 大小
        XSSFFont font = (XSSFFont) workbook.createFont();
        if (fontColor != null) {
            font.setColor(fromStrToARGB(fontColor));
        }
        cellStyle.setFont(font);
        //设置精度
        if (pr != null) {
            cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(pr)); //例如   p = 0.00
        }
        if (pe != null) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(pe));
        }
        dataCell.setCellStyle(cellStyle);
    }

    public static void setHeadStyleOfFormula(ExportContext context, Element node, SXSSFCell dataCell) {
        SXSSFWorkbook workbook = context.getWorkbook();
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        //设置边框
        setAllBorderStyle(cellStyle);
        //设置 背景 字体颜色
        String fillColor = node.attributeValue(ExcelConstants.HEAD_FILLCOLOR);//表头  背景颜色
        String fontColor = node.attributeValue(ExcelConstants.HEAD_FONTCOLOR);//表头  字体颜色
        //背景景颜色
        if (fillColor != null) {
            cellStyle.setFillForegroundColor(fromStrToARGB(fillColor));
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        //设置字体 颜色 大小
        XSSFFont font = (XSSFFont) workbook.createFont();
        if (fontColor != null) {
            font.setColor(fromStrToARGB(fontColor));
        }
        cellStyle.setFont(font);
        dataCell.setCellStyle(cellStyle);
    }

    private static void setAllBorderStyle(XSSFCellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        cellStyle.setAlignment(HorizontalAlignment.CENTER); // 居中
    }

    /**
     * 填充数据 设置样式
     *
     * @param context
     * @param key     xml节点名称
     * @param value   数据值
     * @param cell    单元格对象
     */
    public static void cellFillData(ExportContext context, String key, Object value, SXSSFCell cell) {
        if (key != null) {
            String s = value.toString();
            boolean matches = pattern.matcher(value.toString()).matches();
            if (matches) {
                cell.setCellValue(Double.parseDouble(s));
            } else {
                cell.setCellValue(s);
            }
        }
        System.out.println("行_列_value:" + cell.getRowIndex() + "_" + cell.getColumnIndex() + ":" + value);
        //设置通用样式
        cell.setCellStyle(null);
        //满足条件 改变单元格颜色或者字体颜色
        changeCellStyle(context, key, value, cell);
    }

    /**
     * 处理高亮  字体变色
     *
     * @param context
     * @param key
     * @param valueObj
     * @param cell
     */
    private static void changeCellStyle(ExportContext context, String key, Object valueObj, SXSSFCell cell) {
        Map<String, IfCondition> ifMap = context.getIfMap();
        IfCondition ifCondition = ifMap.get(key);
        if (ifCondition != null) {
            String value = valueObj == null ? null : valueObj.toString();
            String symbol = ifCondition.getIfSymbol();
            String target = ifCondition.getIfTarget();
            XSSFCellStyle cellStyle = ifCondition.getCellStyle();
            boolean flag = false;
            assert value != null;
            boolean matches = pattern.matcher(value).matches();
            if (matches) {
                //如果是比较 目标中含有% 则先去除%
                if (target.contains("%")) {
                    value = value.replaceAll("%", "");
                    target = target.replaceAll("%", "");
                }
                BigDecimal valueB = new BigDecimal(value);
                BigDecimal targetB = new BigDecimal(target);
                // 1  0  -1  大于  等于  小于
                int i = valueB.compareTo(targetB);
                switch (symbol) {
                    case "=":
                        if (i == 0) {
                            flag = true;
                        }
                        break;
                    case ">":
                        if (i > 0) {
                            flag = true;
                        }
                        break;
                    case ">=":
                        if (i >= 0) {
                            flag = true;
                        }
                        break;
                    case "<":
                        if (i < 0) {
                            flag = true;
                        }
                        break;
                    case "<=":
                        if (i <= 0) {
                            flag = true;
                        }
                        break;
                    default:
                        System.out.println("无效的符号");
                        break;
                }
            } else {//字符串判断
                switch (symbol) {
                    case "=":
                        if (Objects.equals(value, target)) {
                            flag = true;
                        }
                    case "!=":
                        if (Objects.equals(value, target)) {
                            flag = true;
                        }
                    default:
                        System.out.println("无效的符号");
                }
            }
            if (flag) {
                cell.setCellStyle(cellStyle);
            }
        }
    }

    /**
     * 额外列:指 动态表头模式中,非动态列(除第一行),这种数据需要单独查出再根据第一列对应写入
     * <p>
     * 动态表头模式
     * 创建表头 同时插入数据
     */
    public static ExportContext createHeadAndFillData(ExportContext context, Integer headRowStart, Integer headRowEnd, SXSSFSheet sheet, Element structureElement, Map<String, Object> map) {
        String firstHead = context.getFirstHead();
        String flexHead = context.getFlexHead();
        Integer headUsedRow = context.getHeadUsedRow();
        Integer dataUsedRow = context.getDataUsedRow();
        Integer headUsedCol = context.getHeadUsedCol();
        Integer readCount = context.getReadCount();

        Iterator iterator = structureElement.elementIterator();
        SXSSFRow headRow;
        if (headRowStart <= headUsedRow) {//不能重复create 同一行
            headRow = sheet.getRow(headRowStart);
        } else {
            headRow = sheet.createRow(headRowStart);
            context.setHeadUsedRow(headRowStart);
        }
        while (iterator.hasNext()) {
            Element node = (Element) iterator.next();
            String key = node.getName();//xml  节点名称   字段名  <key></key>
            int size = node.elements().size();//子节点 数量  决定表头的合并单元格的列数
            String parentNodeName = node.getParent().getName();// 父节点  名称
            String head_name = node.attributeValue(ExcelConstants.HEAD_NAME);
            boolean needHeadUsedColAdd = false;
            //不含子表头
            Map<String, Integer> indexList = context.getIndexList();
            if (!node.hasMixedContent()) {
                //1. 设置表头-----------------------------------------------------
                SXSSFCell headCell;
                //!!!!!!!这个判断的目的是, 类似第一列的这种非动态列不能重复创建的表头(靠其他部分逻辑填充数据)!!!!!
                if (!(Objects.equals(parentNodeName, ExcelConstants.TABLE_STRUCTURE) && readCount > 0)) {//不是第一条数据  且 此节点没有子节点 且 没有父节点  则不根据节点创建head
                    //处理 条件 等式 不等式样式
                    doIfCondition(context, node, key);
                    headCell = headRow.createCell(headUsedCol);
                    //设置表头样式
                    CellStyle cellStyle = initSetHeadCellStyle(context, node, key);
                    headCell.setCellStyle(cellStyle);
                    //设置表头名称
                    headCell.setCellValue(head_name);
                    //记录每一列的位置  除了第一列(第一列记录的是行数)
                    if (!Objects.equals(key, firstHead)) {
                        // 父节点值_此节点名称   例如  musk_phone
                        String filesKey = flexHead + "_" + map.get(flexHead) + "_" + key; //写入数据在哪一行写 只需根据第一列字段判断
                        indexList.put(filesKey, headUsedCol);
                    }
                    //合并表头单元格(行)
                    if (headRowEnd > 0 & headRowEnd > headRowStart) {//超过一行
                        CellRangeAddress region = new CellRangeAddress(headRowStart, headRowEnd, headUsedCol, headUsedCol);
                        sheet.addMergedRegion(region);
                    }
                    needHeadUsedColAdd = true;
                    //记录额外列
                    indexList.put(key, headUsedCol);
                }
                //2. 正文  填入数据--------------------------------------------------------
                String firstHeadRowCountKey = firstHead + "_" + map.get(firstHead); //写入数据在哪一行写 只需根据第一列字段判断
                SXSSFRow dataRow;
                if (indexList.get(firstHeadRowCountKey) != null) {
                    dataRow = sheet.getRow(indexList.get(firstHeadRowCountKey));//!!!!!这里不能用createRow 否则会清空之前本行的数据
                } else {
                    context.setDataUsedRow(++dataUsedRow);
                    dataRow = sheet.createRow(dataUsedRow);
                    indexList.put(firstHeadRowCountKey, dataUsedRow);
                    SXSSFCell cell = dataRow.createCell(0);
                    String value = map.get(firstHead).toString();
                    cellFillData(context, firstHead, value, cell);//填入首列 数据
                }
                SXSSFCell dataCell = dataRow.createCell(headUsedCol);
                String value = map.get(key) == null ? "" : map.get(key).toString();
                cellFillData(context, key, value, dataCell);
                if (needHeadUsedColAdd) {
                    context.setHeadUsedCol(++headUsedCol);//下一个字段
                }
                context.setIndexList(indexList);
            } else {//递归-------------------------------------------------
                //处理 条件 等式 不等式样式
                doIfCondition(context, node, key);
                //设置表头文字
                SXSSFCell headCell = headRow.createCell(headUsedCol);
                //设置表头样式
                CellStyle cellStyle = initSetHeadCellStyle(context, node, key);
                headCell.setCellStyle(cellStyle);
                //设置表头名称
                String value = map.get(node.getName()).toString();
                headCell.setCellValue(value);
                //记录表头起始位置
                String flexFilesKey = key + "_" + map.get(key); //写入数据在哪一行写 只需根据第一列字段判断 flexFiles记录列
                indexList.put(flexFilesKey, headUsedCol);
                context.setIndexList(indexList);
                //合并单元格  (列)
                CellRangeAddress region = new CellRangeAddress(headRowStart, headRowStart, headUsedCol, headUsedCol + size - 1);
                sheet.addMergedRegion(region);
                headRowStart++;//子表头 行+1
                createHeadAndFillData(context, headRowStart, headRowEnd, sheet, node, map);
            }
        }
        return context;
    }

    /**
     * 固定表头模式  创建表头   添加样式集合
     *
     * @param context
     * @param headRowStart     表头开始行  从0开始  递归后变化
     * @param headRowEnd       表头结束行  xml 指定
     * @param sheet
     * @param structureElement 表结构dom对象
     */
    public static ExportContext createFixedHead(ExportContext context, Integer headRowStart, Integer headRowEnd, SXSSFSheet sheet, Element structureElement) {
        Integer headUsedRow = context.getHeadUsedRow();
        Integer headUsedCol = context.getHeadUsedCol();
        Iterator iterator = structureElement.elementIterator();
        SXSSFRow headRow;
        if (headRowStart <= headUsedRow) {//不能重复create 同一行
            headRow = sheet.getRow(headRowStart);
        } else {
            headRow = sheet.createRow(headRowStart);
            context.setHeadUsedRow(headRowStart);
        }
        while (iterator.hasNext()) {
            Element node = (Element) iterator.next();
            String key = node.getName();//xml  节点名称   字段名  <key></key>
            int size = node.elements().size();//子节点 数量  决定表头的合并单元格的列数
            String head_name = node.attributeValue(ExcelConstants.HEAD_NAME);
            Map<String, Integer> indexList = context.getIndexList();
            if (!node.hasMixedContent()) {
                //处理 条件 等式 不等式样式
                doIfCondition(context, node, key);
                //设置表头-----------------------------------------------------
                SXSSFCell headCell;
                headCell = headRow.createCell(headUsedCol);
                //设置表头样式
                CellStyle cellStyle = initSetHeadCellStyle(context, node, key);
                headCell.setCellStyle(cellStyle);
                //设置表头名称
                headCell.setCellValue(head_name);
                //记录每一列的位置  除了第一列(第一列记录的是行数)
                indexList.put(key, headUsedCol);//写入数据在哪一行写 只需根据第一列字段判断
                context.setIndexList(indexList);
                //合并表头单元格(行)
                if (headRowEnd > 0 & headRowEnd > headRowStart) {//超过一行
                    CellRangeAddress region = new CellRangeAddress(headRowStart, headRowEnd, headUsedCol, headUsedCol);
                    sheet.addMergedRegion(region);
                    setMergedRegionStyle(sheet, region);
                }
                context.setHeadUsedCol(++headUsedCol);//下一个字段
            } else {//递归-------------------------------------------------
                //处理 条件 等式 不等式样式
                doIfCondition(context, node, key);
                //设置表头文字
                SXSSFCell headCell = headRow.createCell(headUsedCol);
                //设置表头样式
                CellStyle cellStyle = initSetHeadCellStyle(context, node, key);
                headCell.setCellStyle(cellStyle);
                //设置表头名称
                headCell.setCellValue(head_name);
                //合并单元格  (列)
                CellRangeAddress region = new CellRangeAddress(headRowStart, headRowStart, headUsedCol, headUsedCol + size - 1);
                sheet.addMergedRegion(region);
                headRowStart++;//子表头 行+1
                createFixedHead(context, headRowStart, headRowEnd, sheet, node);
            }
        }
        return context;
    }


    /**
     * 设置样式
     *
     * @param context
     * @param node
     * @param key     动态表头 和 固定表头的  传入 key 都直接用字段名做做key即可
     *                用到了两个强转类 便于直接使用rgb颜色
     */
    private static CellStyle initSetHeadCellStyle(ExportContext context, Element node, String key) {
        SXSSFWorkbook workbook = context.getWorkbook();
        XSSFCellStyle headCellStyle = (XSSFCellStyle) workbook.createCellStyle();
        String fillColor = node.attributeValue(ExcelConstants.HEAD_FILLCOLOR);
        String fontColor = node.attributeValue(ExcelConstants.HEAD_FONTCOLOR);
        //设置 边框  背景
        setAllBorderStyle(headCellStyle);
        //设置填充颜色
        if (fillColor != null) {
            //背景景颜色
            headCellStyle.setFillForegroundColor(fromStrToARGB(fillColor));
            headCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        //设置字体 颜色 大小
        XSSFFont font = (XSSFFont) workbook.createFont();
        if (fontColor != null) {
            font.setColor(fromStrToARGB(fontColor));
        }
        font.setBold(true);
        headCellStyle.setFont(font);
        return headCellStyle;
    }

    /**
     * 对所有没有边框的单元格设置边框   不支持粗细自定义
     *
     * @param context
     */
    public static void setDataBorderStyle(ExportContext context) {
        SXSSFWorkbook workbook = context.getWorkbook();
        XSSFCellStyle dataCellStyle = (XSSFCellStyle) workbook.createCellStyle();
        Integer headUsedCol = context.getHeadUsedCol();
        Integer dataUsedRow = context.getDataUsedRow();
        Integer headUsedRow = context.getHeadUsedRow();
        SXSSFSheet sheet = workbook.getSheet("名单分配");
        //设置 边框
        setAllBorderStyle(dataCellStyle);
        // TODO: 2019/11/11  大后期再优化吧  需要在写数据前知道 行和列的最大值
        for (int j = headUsedRow + 1; j <= dataUsedRow; j++) {
            SXSSFRow row = sheet.getRow(j);
            for (int i = 0; i < headUsedCol; i++) {
                SXSSFCell cell = row.getCell(i);
                if (cell == null) {
                    cell = row.createCell(i);
                }
                CellStyle cellStyle = cell.getCellStyle();
                BorderStyle borderRightEnum = cellStyle.getBorderRightEnum();
                if (borderRightEnum == BorderStyle.NONE) {//没有设置有边框  即此单元格没有用设置边
                    cell.setCellStyle(dataCellStyle);
                }
            }
        }
    }

    /**
     * 处理条件判断高亮 字体变色: 例如   该列数值大于5 字体变成红色 背景变成蓝色
     *
     * @param context
     * @param node
     * @param key
     */
    private static ExportContext doIfCondition(ExportContext context, Element node, String key) {
        SXSSFWorkbook workbook = context.getWorkbook();
        Map<String, IfCondition> ifMap = context.getIfMap();
        String ifSymbol = node.attributeValue(ExcelConstants.HEAD_IFSYMBOL);
        String ifTarget = node.attributeValue(ExcelConstants.HEAD_IFTARGET);
        if (ifSymbol != null) {
            // TODO: 2019/10/22  后期补全校验   只执行允许的符号
            XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
            //设置边框
            setAllBorderStyle(cellStyle);
            IfCondition condition = new IfCondition();
            if (node.attributeValue(ExcelConstants.HEAD_CHANGEDATAFILLCOLOR) != null) {
                String fillColor = node.attributeValue(ExcelConstants.HEAD_CHANGEDATAFILLCOLOR);
                cellStyle.setFillForegroundColor(fromStrToARGB(fillColor));
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            if (node.attributeValue(ExcelConstants.HEAD_CHANGEDATAFONTCOLOR) != null) {
                String fontColor = node.attributeValue(ExcelConstants.HEAD_CHANGEDATAFONTCOLOR);
                XSSFFont font = (XSSFFont) workbook.createFont();
                font.setColor(fromStrToARGB(fontColor));
                cellStyle.setFont(font);
            }
            condition.setIfTarget(ifTarget);
            condition.setIfSymbol(ifSymbol);
            condition.setCellStyle(cellStyle);
            ifMap.put(key, condition);
            context.setIfMap(ifMap);
        }
        return context;
    }

    //合并后的单元格设置样式 和 未合并的设置样式有区别
    private static void setMergedRegionStyle(Sheet sheet, CellRangeAddress region) {
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
    }

    //十六进制色码转成poi 可用颜色对象
    private static XSSFColor fromStrToARGB(String hex) {
        int red = Integer.valueOf(hex.substring(1, 3), 16);
        int green = Integer.valueOf(hex.substring(3, 5), 16);
        int blue = Integer.valueOf(hex.substring(5, 7), 16);
        return new XSSFColor(new java.awt.Color(red, green, blue));
    }

    /**
     * @param tmplateCode 指定模板的根节点名称
     * @return
     */
    public static Element getElement(String tmplateCode) {
        //获取Document对象
        SAXReader reader = new SAXReader();
        Document document = null;
        try {
            document = reader.read(ResourceUtils.getFile("classpath:excelConfig.xml"));
        } catch (Exception e) {
            logger.error("加载报表配置文件异常");
        }
        assert document != null;
        Element rootElement = document.getRootElement();//整个xml 对象
        if (document.getRootElement() != null && rootElement.element(tmplateCode) == null) {
            logger.error("没有" + tmplateCode + "的相关配置");
            // TODO: 2019/10/30  throw
        }
        return rootElement.element(tmplateCode);
    }

    /**
     * 填充额外字段数据
     *
     * @param context
     * @param dataListext
     * @param sheet
     */
    public static void setExtData(ExportContext context, List<Map<String, Object>> dataListext, SXSSFSheet sheet) {
        String firstHead = context.getFirstHead();
        Map<String, Integer> indexList = context.getIndexList();
        for (int i = 0; i < dataListext.size(); i++) {
            System.out.println("+++++++++++++++++++++++++++++++++第" + i + "条数据");
            Map<String, Object> map = dataListext.get(i);
            //行列 都存在     直接填写数据
            for (Map.Entry<String, Object> entry : map.entrySet()) {
                if (!Objects.equals(entry.getKey(), firstHead)) { //第一列已经填入数据 步用管
                    String rowKey = firstHead + "_" + map.get(firstHead); //行
                    Integer rowId = indexList.get(rowKey);
                    Integer colId = indexList.get(entry.getKey());
                    SXSSFRow extRow = sheet.getRow(rowId);
                    if (colId != null && rowId != null) {
                        SXSSFCell cell = extRow.createCell(colId);
                        ExcelExportUtil.cellFillData(context, entry.getKey(), entry.getValue(), cell);
                    }
                }
            }
        }
    }
}
