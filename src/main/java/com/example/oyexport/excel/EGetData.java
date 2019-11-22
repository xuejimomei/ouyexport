package com.example.oyexport.excel;

import com.example.oyexport.common.ExcelConstants;
import com.example.oyexport.common.SpringContext;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public enum EGetData {
    MYSQL(ExcelConstants.DATA_HEAD_DB_MYSQL) {
        @Override
        public List<Map<String, Object>> queryData(String mapperStr, String methodName, Map map) {
            Object mapper = SpringContext.getBean(mapperStr);
            List<Map<String, Object>> list = null;
            Class<?> aClass = mapper.getClass();
            try {
                Method method = aClass.getMethod(methodName, Map.class);
                list = (List<Map<String, Object>>) method.invoke(mapper, map);
            } catch (Exception e) {
                e.printStackTrace();
            }
            return list;
        }
    }, MONGODB(ExcelConstants.DATA_HEAD_DB_MONGODB) {
        @Override
        public List<Map<String, Object>> queryData(String mapperStr, String methodName, Map map) {
            //TODO 暂时没有mongo需求
            return null;
        }
    };
    // TODO: 2019/10/31 后期的跟进要求
    private String type;

    EGetData(String t) {
        this.type = t;
    }

    public static EGetData getType(String t) {
        for (EGetData c : EGetData.values()) {
            if (Objects.equals(c.getType(), t)) {
                return c;
            }
        }
        return null;
    }

    public String getType() {
        return type;
    }

    public abstract List<Map<String, Object>> queryData(String mapperStr, String methodName, Map map);
}
