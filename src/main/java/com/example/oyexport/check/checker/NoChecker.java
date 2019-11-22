package com.example.oyexport.check.checker;

import com.example.oyexport.common.MyException;
import com.google.common.primitives.Longs;

/**
 * 数字校验注释处理
 */
public class NoChecker extends BaseChecker {
    public static final String key = "NoChecker";

    @Override
    public void check(Object value) {
        Long l = Longs.tryParse(value.toString());
        if (l == null) {
            throw new MyException("呵呵呵", "参数不对");
        }
    }
}
