package com.example.oyexport.vo;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.example.oyexport.check.annotation.Base;
import com.example.oyexport.check.annotation.No;
import com.example.oyexport.check.checker.CheckUtil;
import com.example.oyexport.common.ExcelConstants;
import com.example.oyexport.common.ExcelExportUtil;
import com.example.oyexport.common.MyException;
import com.google.common.base.Strings;
import org.dom4j.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

public class ExcelExportInVO {
    private static final Logger logger = LoggerFactory.getLogger(ExcelExportInVO.class);
    @Base
    private String paramJson; //json 入参
    private F00000 vo; //json 校验转化后的vo对象

    public void check() throws Exception {
        try {
            JSON.parse(paramJson);
        } catch (Exception e) {
            throw new MyException("0001", "json格式错了");
        }
        JSONObject jsonObject = JSON.parseObject(paramJson);
        String tmplateCode = jsonObject.getString(ExcelConstants.VO_TMPLATECODE);
        if (Strings.isNullOrEmpty(tmplateCode)) {
            throw new MyException("0002", "tmplateCode不能为空");
        } else if (Objects.equals(tmplateCode, ExcelConstants.VO_F00000)) {
            throw new MyException("0002", "tmplateCode不能为F00000");
        }
        String aPackage = this.getClass().getPackage().getName();
        Class<?> aClass;
        try {
            aClass = Class.forName(aPackage + "." + tmplateCode);
        } catch (ClassNotFoundException e) {
            aClass = Class.forName(aPackage + "." + ExcelConstants.VO_F00000);//没有对应类  则使用F00000
        }
        Object object = JSONObject.parseObject(paramJson, aClass);
        F00000 epVO = (F00000) object;
        //vo 中设置 xml 解析的信息
        Element el = ExcelExportUtil.getElement(tmplateCode);
        String type = el.element(ExcelConstants.TABLE_TYPE).getText();
        epVO.setTmplateCode(tmplateCode);
        if (epVO.getFilesName() == null) {
            String filesName = el.element(ExcelConstants.TABLE_FILESNAME).getText();
            epVO.setFilesName(filesName);
        }
        if (epVO.getTableName() == null) {
            String tableName = el.element(ExcelConstants.TABLE_TABLENAME).getText();
            epVO.setTableName(tableName);
        }
        //校验入参
        validate(epVO);
        //校验成功 设置vo
        this.vo = epVO;
    }

    public String getParamJson() {
        return paramJson;
    }

    public void setParamJson(String paramJson) {
        this.paramJson = paramJson;
    }

    public F00000 getVo() {
        return vo;
    }

    public void setVo(F00000 vo) {
        this.vo = vo;
    }

    public final void validate(F00000 epVO) throws Exception {
        Class<? extends F00000> aClass = epVO.getClass();
        //获取子类，父类所有属性
        List<Field> fs = new ArrayList<>();
        for (Class<?> clazz = aClass; clazz != Object.class; clazz = clazz.getSuperclass()) {
            Field[] fieldArr = clazz.getDeclaredFields();
            fs.addAll(Arrays.asList(fieldArr));
        }
        for (Field f : fs) {
            try {
                f.setAccessible(true);
                Annotation[] as = f.getAnnotations();
                for (Annotation a : as) {
                    if (a instanceof No) {
                        No no = (No) a;
                        CheckUtil.checkers.get(no.key()).check(f.get(epVO), no.allowNull());
                    } else if (a instanceof Base) {
                        Base b = (Base) a;
                        CheckUtil.checkers.get(b.key()).check(f.get(epVO), b.allowNull());
                    }
                }
            } catch (Throwable e) {
                if (e instanceof MyException) {
                    throw e;
                }
                logger.error("校验参数异常", e);
            }
        }
    }
}
