package com.example.oyexport.vo;

import com.example.oyexport.check.annotation.Base;

public class F00000<T> {
    @Base
    private String tmplateCode;//tmplateCode 为必填参数
    private String filesName;//可根据前端传参自定义,前端不传参则按照xml 配置的
    private String tableName;//可根据前端传参自定义,前端不传参则按照xml 配置的

    private Integer pageNo =0;
    private Integer pageSize=50;
    public String getTmplateCode() {
        return tmplateCode;
    }

    public void setTmplateCode(String tmplateCode) {
        this.tmplateCode = tmplateCode;
    }

    public String getFilesName() {
        return filesName;
    }

    public void setFilesName(String filesName) {
        this.filesName = filesName;
    }

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    public Integer getPageNo() {
        return pageNo;
    }

    public void setPageNo(Integer pageNo) {
        this.pageNo = pageNo;
    }

    public Integer getPageSize() {
        return pageSize;
    }

    public void setPageSize(Integer pageSize) {
        this.pageSize = pageSize;
    }
}
