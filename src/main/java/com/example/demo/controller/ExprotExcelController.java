package com.example.demo.controller;

import com.example.demo.utils.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by xkx on 2019/4/14.
 */
@Controller
@EnableAutoConfiguration
public class ExprotExcelController {

    /**
     * 导出报表
     */
    @RequestMapping(value = "/export")
    public void export(HttpServletRequest request,HttpServletResponse response) throws Exception {
        //获取数据
        List<Map> list = new ArrayList<>();
        for (int i=0 ; i<50 ; i++ ) {
            Map map = new HashMap();
            map.put("mingcheng", "名称_" + (i+1));
            map.put("sex", i % 2 == 0 ? "男" : "女");
            map.put("age", "年龄_" + (i+1) + "");
            map.put("school", "学校_" + (i+1) + "");
            map.put("banji", "班级_" + (i+1) + "");
            list.add(map);
        }

        //excel标题
        String[] title = {"名称","性别","年龄","学校","班级"};

         //excel文件名
        String fileName = "学生信息表"+System.currentTimeMillis()+".xls";
        //sheet名
        String sheetName = "学生信息表";
        String [][] content = new String[list.size()][];
        for (int i = 0; i < list.size(); i++) {
            content[i] = new String[title.length];
            Map obj = list.get(i);
            content[i][0] = obj.get("mingcheng").toString();
            content[i][1] = obj.get("sex").toString();
            content[i][2] = obj.get("age").toString();
            content[i][3] = obj.get("school").toString();
            content[i][4] = obj.get("banji").toString();
        }

        //创建HSSFWorkbook
        HSSFWorkbook wb = ExcelUtil.getHSSFWorkbook(sheetName, title, content, null);

        //响应到客户端
        try {
            this.setResponseHeader(response, fileName);
            OutputStream os = response.getOutputStream();
            wb.write(os);
            os.flush();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //发送响应流方法
    public void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            try {
                fileName = new String(fileName.getBytes(),"ISO8859-1");
            } catch (UnsupportedEncodingException e) {
                e.printStackTrace();
            }
            response.setContentType("application/octet-stream;charset=ISO8859-1");
            response.setHeader("Content-Disposition", "attachment;filename="+ fileName);
            response.addHeader("Pargam", "no-cache");
            response.addHeader("Cache-Control", "no-cache");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}
