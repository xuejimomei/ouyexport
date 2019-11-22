package com.example.oyexport.common;

public class ExcelConstants {
    //vo
    public static final String VO_F00000 = "F00000";        // 文件名称
    public static final String VO_TMPLATECODE = "tmplateCode";        // 文件名称

    //table 属性
    public static final String TABLE_FILESNAME = "filesName";        // 文件名称
    public static final String TABLE_TABLENAME = "tableName";        // 表名称
    public static final String TABLE_SHEETNAME = "sheetName";        // sheet 名称
    public static final String TABLE_TYPE = "type";                  // 类型:  1  2  分别表示 fixedColunm   和 flexColunm    是否动态表头
    public static final Integer TABLE_TYPE_1 = 1;                  // 类型:  1  不含动态表头
    public static final Integer TABLE_TYPE_2 = 2;                  // 类型:  2  有态表头
    public static final String TABLE_HEADROWCOUNT = "headRowCount"; // 表头的总行数  懒得遍历计算了
    public static final String TABLE_FIRSTHEAD = "firstHead";        // 第一列       懒得遍历计算了
    public static final String TABLE_FLEXHEAD = "flexHead";         // 动态列       懒得遍历计算了
    public static final String TABLE_STRUCTURE = "structure";        // 表结构
    public static final String TABLE_STRUCTUREEXT = "structureExt";  // 额外结构 比如:合计
    public static final String TABLE_DATA = "data";  // 额外结构 比如:合计
    //head 属性
    public static final String HEAD_NAME = "name";                  // 表头名称(翻译)  必填
    public static final String HEAD_FILLCOLOR = "fillColor";        // 填充颜色
    public static final String HEAD_FONTCOLOR = "fontColor";        // 字体颜色
    public static final String HEAD_IFSYMBOL = "ifSymbol";       // 判断条件符号   数字: > <  = >=  <=   字符串: =   !=
    public static final String HEAD_IFTARGET = "ifTarget";       // 判断条件比较对象
    public static final String HEAD_LT = "lt";                      // 小于
    public static final String HEAD_LTE = "lte";                    // 小于等于
    public static final String HEAD_GT = "gt";                      // 大于
    public static final String HEAD_GTE = "gte";                    // 大于等于
    public static final String HEAD_CHANGEDATAFONTCOLOR = "changeDataFontColor";  // 大于小于  满足条件  设置 数据字体颜色
    public static final String HEAD_CHANGEDATAFILLCOLOR = "changeDataFillColor";  // 大于小于  满足条件  设置 数据填充颜色
    public static final String HEAD_DATAFONTCOLOR = "dataFontColor";  //  数据字体默认颜色
    public static final String HEAD_DATAFILLCOLOR = "dataFillColor";  //  数据填充默认颜色
    public static final String HEAD_PRECISION = "precision";                      // 总计  平均值  精度 0.00
    public static final String HEAD_PERCENT = "percent";                          // 总计  平均值  百分数 0.00%
    //structureExt   总计  平均值 .....
    public static final String EXT_HEAD_TOTAL = "total";            // 总计
    public static final String EXT_HEAD_AVERAGE = "average";            // 总计
    //data   数据源信息
    public static final String DATA_METHOD = "method";            // 方法
    public static final String DATA_METHODEXT = "methodExt";            // 方法
    //data  属性
    public static final String DATA_HEAD_MAPPER = "mapper";            // mapper
    public static final String DATA_HEAD_DB = "db";            // 数据库
    public static final String DATA_HEAD_DB_MYSQL = "MYSQL";            // 数据库
    public static final String DATA_HEAD_DB_MONGODB = "MONGODB";            // 数据库

}