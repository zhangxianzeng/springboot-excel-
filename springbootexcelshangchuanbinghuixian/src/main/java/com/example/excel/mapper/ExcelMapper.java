package com.example.excel.mapper;


import com.example.excel.entity.Excels;
import org.apache.ibatis.annotations.Insert;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Select;
import java.util.List;

@Mapper
public interface ExcelMapper {
    @Insert("INSERT INTO excelDemo (num,names,links,passwords,filename) VALUES (#{num},#{names},#{links},#{passwords},#{filename})")
    int insert(Excels excels);
    @Select("select * from excelDemo")
    List<Excels> select();
    @Select("select num,names,links,passwords from excelDemo where filename=#{filename}")
    List<Excels> select1(String filename);

    @Select("select distinct filename from excelDemo ")
    List<Excels> select2();

}
