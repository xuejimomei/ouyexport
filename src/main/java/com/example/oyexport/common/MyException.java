package com.example.oyexport.common;

public class MyException extends RuntimeException {

    private String retCd;  //异常对应的返回码
    private String msgDes;  //异常对应的描述信息

    public MyException() {
        super();
    }

    public MyException(String message) {
        super(message);
        msgDes = message;
    }

    public MyException(String retCd, String msgDes) {
        super();
        this.retCd = retCd;
        this.msgDes = msgDes;
    }

    public String getRetCd() {
        return retCd;
    }

    public String getMsgDes() {
        return msgDes;
    }
}