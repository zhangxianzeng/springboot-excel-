package com.example.excel.controller;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import com.example.excel.entity.Excels;
import com.example.excel.service.ExcelsService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.ibatis.annotations.Mapper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;

import java.net.URLEncoder;
import java.nio.charset.Charset;
import java.util.*;

@Api(value = "/",description = "这是所有的接口文档")
@RestController
@RequestMapping("/act")
public class ExcelController {
    @Value("${file-save-path}")
    private String filePath;
    @Autowired
    private ExcelsService excelsService;

    @ApiOperation(value = "添加文件", httpMethod = "POST")
    @PostMapping("/insert")
    public Map excelInsert(HttpServletRequest request, @RequestParam MultipartFile file) {
        Map map = new HashMap();
        InputStream is = null;
        try {
            File dir = new File(filePath);
            if (!dir.exists()) {
                dir.mkdirs();
            }
            String newFileName = file.getOriginalFilename();
            File newFile = new File(filePath + newFileName);
            //复制操作
            file.transferTo(newFile);
            //读取保存的文件
            is = new FileInputStream((filePath + newFileName));
            System.out.println(newFile.getName());
            Workbook wb = null;
            if ((filePath + newFileName).endsWith(".xlsx")) {
                wb = new XSSFWorkbook(is);
                //HSSFWorkbook不能为xls或者et结尾，如果想要那么用XSSFWorkbook表示以什么结尾的
            } else if ((filePath + newFileName).endsWith(".xls") || (filePath + newFileName).endsWith(".et")) {
                wb = new HSSFWorkbook(is);
            }
            /* 读EXCEL文字内容 */
            // 获取第一个sheet表，也可使用sheet表名获取
            Sheet sheet = wb.getSheetAt(0);
            List<Map<String, String>> sheetList = new ArrayList<Map<String, String>>();
            List<String> titles = new ArrayList<String>();
            int rowSize = sheet.getLastRowNum() + 1;
            for (int j = 0; j < rowSize; j++) {
                Row row = sheet.getRow(j);
                if (row == null) {
                    continue;
                }
                int cellSize = row.getLastCellNum();
                if (j == 0) {
                    for (int k = 0; k < cellSize; k++) {
                        Cell cell = row.getCell(k);
                        titles.add(cell.toString());
                    }
                } else {
                    Map<String, String> rowMap = new HashMap<String, String>();
                    for (int k = 0; k < titles.size(); k++) {
                        Cell cell = row.getCell(k);
                        String key = titles.get(k);
                        String value = null;
                        if (cell != null) {
                            value = cell.toString();
                        }
                        rowMap.put(key, value);
                    }
                    sheetList.add(rowMap);
                    //rowMap.clear();
                }
            }
            wb.close();
            is.close();
            titles.clear();
            String num;
            String name;
            String links;
            String password;
            String filename;
            for (Map<String, String> params : sheetList) {
                num = params.get("序号");
                name = params.get("姓名");
                links = params.get("链接");
                password = params.get("密码");
                Excels excels = new Excels();
                excels.setNum(num);
                excels.setNames(name);
                excels.setLinks(links);
                excels.setPasswords(password);

                excels.setFilename(newFileName);

                int insert = excelsService.insert(excels);
                if (insert > 0) {
                    map.put("code", "00000");
                    map.put("msg", "ok");
                }

            }
        } catch (Exception ex) {
            ex.printStackTrace();
            map.put("code", "-1");
            map.put("msg", "eorry");
        }
        return map;
    }

    @ApiOperation(value = "查询", httpMethod = "POST")
    @PostMapping("/select")
    public Map select() {
        Map map = new HashMap();
        try {
            List<Excels> list = excelsService.select();
            map.put("date", list);
            map.put("code", "00000");
            map.put("msg", "ok");
        } catch (Exception e) {
            e.printStackTrace();
            map.put("code", "-1");
            map.put("msg", "eorry");
        }
        return map;
    }

    @ApiOperation(value = "查询某个文件的数据", httpMethod = "POST")
    @PostMapping("/select1")
    public Map select1(String filname) {
        Map map = new HashMap();
        try {
            List<Excels> list = excelsService.select1(filname);
            map.put("date", list);
            map.put("code", "00000");
            map.put("msg", "ok");
        } catch (Exception e) {
            e.printStackTrace();
            map.put("code", "-1");
            map.put("msg", "eorry");
        }
        return map;
    }

    @ApiOperation(value = "查询数据库中的所有xlsx文件", httpMethod = "POST")
    @PostMapping("/select2")
    public Map select2() {
        Map map = new HashMap();
        try {
            List<Excels> list = excelsService.select2();
            map.put("date", list);
            map.put("code", "00000");
            map.put("msg", "ok");
        } catch (Exception e) {
            e.printStackTrace();
            map.put("code", "-1");
            map.put("msg", "eorry");
        }
        return map;
    }


    //下面的这个是不经过数据库的
    @ApiOperation(value = "查询上传的后的服务器文件夹下所以的文件名称", httpMethod = "POST")
    @PostMapping("/select3")
    public Map select3() {
        File file = new File(filePath);
        Map map = new HashMap();
        File[] fileList = file.listFiles();
        for (File f : fileList) {
            if (f.isFile() && (f.getName().endsWith(".xls") || f.getName().endsWith(".xlsx"))) {
                map.put("date", f.getName());
                map.put("code", "00000");
                map.put("msg", "ok");

            }

        }
        return map;
    }

    @ApiOperation(value = "打开服务器文件夹下某个名称的文件", httpMethod = "POST")
    @PostMapping("/select4")
    public Map select3(String filename) {

        ExcelReader reader = ExcelUtil.getReader(filePath+filename);
        List<Map<String,Object>> readAll = reader.readAll();
        Map map = new HashMap();
        map.put("date", readAll);
        map.put("code", "00000");
        map.put("msg", "ok");
        return map;
    }



    /**
     * 平台实现更新包下载
     *
     * @param request
     * @param response
     * @return
     * @throws UnsupportedEncodingException
     */
    @ApiOperation(value = "下载某个文件", httpMethod = "POST")
    @RequestMapping("/download")
//需要从浏览器进行访问否则打不开
    public Object downloadFile(HttpServletRequest request,
                               HttpServletResponse response, String fileFullName) throws UnsupportedEncodingException {
        //String rootPath = propertiesconfig.getUploadpacketPath();//这里是我在配置文件里面配置的根路径，各位可以更换成自己的路径之后再使用（例如：D：/test）
        String FullPath = filePath + fileFullName;//将文件的统一储存路径和文件名拼接得到文件全路径
        File packetFile = new File(FullPath);
        String fileName = packetFile.getName(); //下载的文件名
        File file = new File(FullPath);
        // 如果文件名存在，则进行下载
        if (file.exists()) {
            // 配置文件下载
            //response.setHeader("Content-Type", "application/octet-stream");
            response.setContentType("application/octet-stream");
            // 下载文件能正常显示中文
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            // 实现文件下载
            byte[] buffer = new byte[1024];
            FileInputStream fis = null;
            BufferedInputStream bis = null;
            try {
                fis = new FileInputStream(file);
                bis = new BufferedInputStream(fis);
                OutputStream os = response.getOutputStream();
                int i = bis.read(buffer);
                while (i != -1) {
                    os.write(buffer, 0, i);
                    i = bis.read(buffer);
                }
                System.out.println("Download the song successfully!");
            } catch (Exception e) {
                System.out.println("Download the song failed!");
            } finally {
                if (bis != null) {
                    try {
                        bis.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if (fis != null) {
                    try {
                        fis.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        } else {//对应文件不存在
            try {
                //设置响应的数据类型是html文本，并且告知浏览器，使用UTF-8 来编码。
                response.setContentType("text/html;charset=UTF-8");

                //String这个类里面， getBytes()方法使用的码表，是UTF-8，  跟tomcat的默认码表没关系。 tomcat iso-8859-1
                String csn = Charset.defaultCharset().name();

                System.out.println("默认的String里面的getBytes方法使用的码表是： " + csn);

                //1. 指定浏览器看这份数据使用的码表
                response.setHeader("Content-Type", "text/html;charset=UTF-8");
                OutputStream os = response.getOutputStream();

                os.write("对应文件不存在".getBytes());
            } catch (IOException e) {
                e.printStackTrace();
            }
//                return R.error("-1","对应文件不存在");
        }
        return null;
    }


    //单个删除某个文件
    @ApiOperation(value = "删除某个文件", httpMethod = "POST")
    @RequestMapping("/deleteFile")
    public Map delFile(String fileFullName) {
        String FullPath = filePath + fileFullName;//将文件的统一储存路径和文件名拼接得到文件全路径
       // File packetFile = new File(FullPath);
        Map map = new HashMap();
       // String fileName = packetFile.getName(); //下载的文件名
        File file = new File(FullPath);
        if (file.exists()){
            file.delete();
            map.put("msg","删除成功");

        }else
        {
            map.put("msg","文件不存在");
        }
        return map;
    }
    @ApiOperation(value = "批量删除某些文件", httpMethod = "POST")
    @RequestMapping("/deleteFilelist")
    public Map delFilelist(String fileFullName) {

        Map map = new HashMap();
        String[] fileFullNames = fileFullName.split(",");
        for (String fileFullName1 : fileFullNames) {
            String FullPath = filePath + fileFullName1;//将文件的统一储存路径和文件名拼接得到文件全路径
            // File packetFile = new File(FullPath);

            // String fileName = packetFile.getName(); //下载的文件名
            File file = new File(FullPath);
            if (file.exists()) {
                file.delete();
                map.put("msg", "删除成功");

            } else {
                map.put("msg", "文件不存在");
            }
        }
        return map;

    }

}
