package com.example.oyexport.controller;

import com.example.oyexport.service.ExcelExportService;
import com.example.oyexport.vo.ExcelExportInVO;
import com.example.oyexport.vo.F00000;
import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("/excel")
public class ExportController {
    private ExcelExportService excelExportService;

    @Autowired
    public ExportController(ExcelExportService excelExportService) {
        this.excelExportService = excelExportService;
    }


    @PostMapping("/export")
    public void report(ExcelExportInVO vo) throws Exception {
        excelExportService.export(vo.getVo());
    }

    //todo 入参json   加校验

    @PostMapping("/data")
    public Page<Object> data(ExcelExportInVO vo) throws Exception {
        F00000 vo1 = vo.getVo();
        Page<Object> p = PageHelper.startPage(vo1.getPageNo(), vo1.getPageSize());
        excelExportService.retunData(vo1);
        int pageNum = p.getPageNum();
        int pages = p.getPages();
        List<Object> result = p.getResult();
        return  p;
    }

}
