package com.example.oyexport.check.checker;


import com.example.oyexport.check.checker.BaseChecker;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public class CheckUtil {
    public static final Map<String, BaseChecker> checkers = Collections.unmodifiableMap(new HashMap<String, BaseChecker>() {{
        put(BaseChecker.key, new BaseChecker());
    }});
}
