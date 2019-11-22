package com.example.oyexport.service;


import com.example.oyexport.common.ExcelConstants;
import com.example.oyexport.common.ExcelExportUtil;
import com.example.oyexport.common.ExportContext;
import com.example.oyexport.excel.EExportType;
import com.example.oyexport.excel.EFormula;
import com.example.oyexport.excel.EGetData;
import com.example.oyexport.vo.F00000;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.dom4j.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.util.*;

@Service
public class ExcelExportService {
    private static final Logger logger = LoggerFactory.getLogger(ExcelExportService.class);


    public void export(F00000 vo) {
        //todo 沟通报表数据量  已沟通 无海量数据, 按年导出即可
        List<Map<String, Object>> list1 = getDataList(vo, ExcelConstants.DATA_METHOD);
        List<Map<String, Object>> list2 = getDataList(vo, ExcelConstants.DATA_METHODEXT);
        ExportContext context = new ExportContext(vo.getTmplateCode());
//        excelExportUtil.export(model,vo, list1, list2);
        export(context, vo, getMaps1(), getMapsext1());
    }


    public List<Map<String, Object>> retunData(F00000 vo) {
        HashMap<String, Object> map = new HashMap<>();
        List<Map<String, Object>> list1 = getDataList(vo, ExcelConstants.DATA_METHOD);
//        List<Map<String, Object>> list2 = getDataList(vo, ExcelConstants.DATA_METHODEXT);
//        map.put("dataList", list1);
//        map.put("dataListExt", list2);
        //处理返回数据
        return list1;
    }


    private List<Map<String, Object>> getDataList(F00000 vo, String dataMethod) {
        Element originElement = ExcelExportUtil.getElement(vo.getTmplateCode());
        Element dataElement = originElement.element(ExcelConstants.TABLE_DATA);
        Element methodElement = dataElement.element(dataMethod);
        Element mapperElement = dataElement.element(ExcelConstants.DATA_HEAD_MAPPER);
        Map<Object, Object> map = new HashMap<>();
        EGetData type = EGetData.getType(methodElement.attributeValue(ExcelConstants.DATA_HEAD_DB));
        assert type != null;
        return type.queryData(mapperElement.getText(), methodElement.getText(), map);
    }

    /**
     * @param context
     * @param vo
     * @param dataList    表格数据
     * @param dataListext 额外数据（数据库中不能在同一条数据中的 信息）
     */
    public void export(ExportContext context, F00000 vo, List<Map<String, Object>> dataList, List<Map<String, Object>> dataListext) {
        String filesName = vo.getFilesName();
        String tableName = vo.getTableName();
        Integer type = context.getType();
        Integer headRowEnd = context.getHeadRowEnd();
        Element structure = context.getStructure();
        Element structureExt = context.getStructureExt();
        SXSSFWorkbook workbook = context.getWorkbook();
        SXSSFSheet sheet = workbook.createSheet(tableName);
        //1. 根据type  类型  创建不同类型的表格  并填充数据
        context = EExportType.getType(type).action(context, dataList, dataListext, headRowEnd, structure, sheet);
        //2. 设置各种函数:目前只包括 总计 平均值
        if (structureExt != null) {
            Iterator iterator = structureExt.elementIterator();
            while (iterator.hasNext()) {
                Element node = (Element) iterator.next();
                //求和   平均值  等.........
                context = EFormula.getType(node.getName()).setFormula(context, headRowEnd, sheet, node);
            }
        }
        //3. 绘制数据单元格边框 (表头 和 总计  平均值..行除)
        ExcelExportUtil.setDataBorderStyle(context);
        try {
//            String outPath = "C:\\Users\\Administrator\\Desktop\\";
            String outPath = "D:\\报表导出目录\\";
//            String fileName = filesName + "_" + LocalDate.now().toString() + ".xlsx";
            String fileName = filesName + "_" + LocalDateTime.now().toString().replace(".", "").replace(":", "").replace(" ", "") + ".xlsx";
            FileOutputStream fos = new FileOutputStream(outPath + fileName);
            workbook.write(fos);
            fos.close();
            //todo
        } catch (Exception e) {
            e.printStackTrace();
            logger.error("报表导出失败");
        }
    }

    private List<Map<String, Object>> getMaps1() {
        List<Map<String, Object>> dataList = new ArrayList<>();
        for (int i = 0; i < 5; i++) {//不同用户不同日期
            Map<String, Object> map = new HashMap<>();
            map.put("report_day", "2019_10_" + i);
            map.put("emp_name", "user" + i);
            map.put("first_follow_rate", i);
            map.put("second_follow_rate", i);
            map.put("thirdly_follow_rate", i);
            dataList.add(map);
        }
        for (int i = 0; i < 5; i++) {//相对第一部分 同日期不同用户
            Map<String, Object> map = new HashMap<>();
            map.put("report_day", "2020_10_" + i);
            map.put("emp_name", "CS" + i);
            map.put("first_follow_rate", i);
            map.put("second_follow_rate", i);
            map.put("thirdly_follow_rate", i);
            dataList.add(map);
        }
        for (int i = 0; i < 5; i++) {//相对第二部分 同用户不同日期
            Map<String, Object> map = new HashMap<>();
            map.put("report_day", "2020_10_" + i);
            map.put("emp_name", "user" + i);
            map.put("first_follow_rate", i);
            map.put("second_follow_rate", i);
            map.put("thirdly_follow_rate", i);
            dataList.add(map);
        }
        return dataList;
    }

    private List<Map<String, Object>> getMapsext1() {
        List<Map<String, Object>> dataList = new ArrayList<>();
        for (int i = 0; i < 5; i++) {//不同用户不同日期
            Map<String, Object> map = new HashMap<>();
            map.put("report_day", "2019_10_" + i);
            map.put("other", i);
            dataList.add(map);
        }
        return dataList;
    }

    private List<Map<String, Object>> getMapsext2() {
        List<Map<String, Object>> dataList = new ArrayList<>();
        for (int i = 0; i < 3; i++) {//不同用户不同日期
            Map<String, Object> map = new HashMap<>();
            map.put("report_day", "2019_10_" + i);
            map.put("other", i);
            dataList.add(map);
        }
        return dataList;
    }

    private List<Map<String, Object>> getMaps2() {
        List<Map<String, Object>> dataList = new ArrayList<>();
        for (int i = 0; i < 5; i++) {//不同用户不同日期
            Map<String, Object> map = new HashMap<>();
            map.put("emp_name", "user" + i);
            map.put("report_day", "2019_10_" + i);
            map.put("first_follow_rate", i);
            map.put("second_follow_rate", i);
            map.put("thirdly_follow_rate", i);
            map.put("other", i);
            dataList.add(map);
        }
        return dataList;
    }


}
