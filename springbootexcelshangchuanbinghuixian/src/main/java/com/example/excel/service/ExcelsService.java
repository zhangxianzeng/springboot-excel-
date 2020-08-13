package com.example.excel.service;


import com.example.excel.entity.Excels;
import com.example.excel.mapper.ExcelMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class ExcelsService {

    @Autowired
    private ExcelMapper excelMapper;

    public int insert(Excels excels) {

        return excelMapper.insert(excels);
    }

    public List<Excels> select() {
        return excelMapper.select();
    }

    public List<Excels> select1(String filename) {
        return excelMapper.select1(filename);
    }
    public List<Excels> select2() {
        return excelMapper.select2();
    }
}
