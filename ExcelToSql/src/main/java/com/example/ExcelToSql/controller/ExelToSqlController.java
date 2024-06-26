package com.example.ExcelToSql.controller;

import com.example.ExcelToSql.service.ExcelToUpdateSqlConverter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ExelToSqlController {

    @Autowired
    ExcelToUpdateSqlConverter excelToUpdateSqlConverter;

    @GetMapping("/testSql")
    public void SqlGenerator(){
        excelToUpdateSqlConverter.ExcelToSqlFileWriter();

    }
}
