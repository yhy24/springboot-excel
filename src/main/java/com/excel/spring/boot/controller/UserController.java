package com.excel.spring.boot.controller;

import com.excel.spring.boot.service.UserService;
import com.excel.spring.boot.util.ExcelUtil;
import com.excel.spring.boot.util.User;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * @Author: yhy
 * @Date: 2018/7/31 10:46
 * @Version 1.0
 */
@RestController
public class UserController {
    @Autowired
    UserService userService;

    @RequestMapping("/get/{id}")
    public String getUser(@PathVariable("id") Integer id) {
        User user = userService.selectUserById(id);
        return user.toString();
    }
    @RequestMapping("/users")
    public String getUsers() {
        List<User> users = userService.selectUsers();
        return users.toString();
    }


    @RequestMapping("/excel")
    public String getExcel(HttpServletRequest request, HttpServletResponse response) {
//        获取数据
//        User user = userService.selectUserById(1);
        List<User> users = userService.selectUsers();
        System.out.println("-------user-------"+users.toString());
//        excel 标题
        String[] titles = {"id","username","password","interest","phone","email","age","mark","code","create_date","modify_date","gender","流水号"};
        SimpleDateFormat smp = new SimpleDateFormat("yyyyMMddHHmmss");
        String format = smp.format(new Date());
//        excel的名称
        String fileName = "user_info-"+format+".xls";
//        sheet 的名字
        String sheetName = "user_info";
        int size = users.size();
        String[][] content = new String[size][];
        for (int i = 0; i < size; i ++) {
            content[i] = new String[titles.length];
            User user = users.get(i);
            content[i][0] = user.getId().toString();
            content[i][1] = user.getUsername();
            content[i][2] = user.getPassword();
            content[i][3] = user.getIntest();
            content[i][4] = user.getPhone();
            content[i][5] = user.getEmail();
            content[i][6] = user.getAge().toString();
            content[i][7] = user.getMark();
            content[i][8] = user.getCode();
            content[i][9] = user.getCreate_date();
            content[i][10] = user.getModify_date();
            content[i][11] = user.getGender();
            content[i][12] = user.getIndicSeq();
        }
        System.out.println("-------content------"+content.toString());
        HSSFWorkbook excel = ExcelUtil.getExcel(sheetName, titles, content, null);
//      相应到客户端
        this.setResponseHeader(response, fileName);
        try {
            OutputStream os = response.getOutputStream();
            excel.write(os);
            os.flush();
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return "Create user_info success";
    }
//发送相应的流
    public void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            fileName = new String(fileName.getBytes(),"ISO8859-1");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        response.setContentType("application/octet-stream;charset=ISO8859-1");
        response.setHeader("Content-Disposition", "attachment;filename="+ fileName);
        response.addHeader("Pargam", "no-cache");
        response.addHeader("Cache-Control", "no-cache");
    }

}
