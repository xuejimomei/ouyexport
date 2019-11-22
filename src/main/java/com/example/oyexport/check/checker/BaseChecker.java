package com.example.oyexport.check.checker;

import com.example.oyexport.common.MyException;
import com.google.common.base.Strings;

/**
 * 非空校验注释处理
 */
public class BaseChecker {
    public static final String key = "BaseChecker";

    public void check(Object value) {
    }

    public void check(Object value, boolean allowNull) {
        if (allowNull) {
            if (value != null && !Strings.isNullOrEmpty(value.toString())) {
                check(value);
            }
        } else {
            if (value == null || Strings.isNullOrEmpty(value.toString())) {
                throw new MyException("呵呵呵", "参数不对");
            } else {
                check(value);
            }
        }
    }
}
