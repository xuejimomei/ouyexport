package com.example.oyexport.common;

import com.example.oyexport.vo.ExcelExportInVO;
import org.aspectj.lang.JoinPoint;
import org.aspectj.lang.annotation.Aspect;
import org.aspectj.lang.annotation.Before;
import org.aspectj.lang.annotation.Pointcut;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

@Aspect
@Component
public class WebLogAspect {
    private static final Logger logger = LoggerFactory.getLogger(WebLogAspect.class);
    private ThreadLocal<Long> startTime = new ThreadLocal<>();

    //定义切入点 该包下的所有函数
    @Pointcut("execution(public * com.example.oyexport.controller..*.*(..))")
    public void webLog() {
        System.out.println("dfsdfsdfasdfasdf");
    }

    @Before("webLog() && args(in)")
    public void doBefore(JoinPoint joinPoint, ExcelExportInVO in) throws Exception {
        in.check();//转换参数后，额外校验
    }

}
