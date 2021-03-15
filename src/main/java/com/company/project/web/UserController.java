package com.company.project.web;

import com.company.project.annotation.CountTime;
import com.company.project.core.Result;
import com.company.project.core.ResultGenerator;
import com.company.project.model.User;
import com.company.project.service.UserService;
import com.company.project.utils.SystemUtils;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

/**
* Created by CodeGenerator on 2021/03/10.
*/
@RestController
@RequestMapping("/user")
public class UserController {
    private static Log logger = LogFactory.getLog(UserController.class);
    @Resource
    private UserService userService;

    @PostMapping("/add")
    public Result add(User user) {
        userService.save(user);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/delete")
    public Result delete(@RequestParam Integer id) {
        userService.deleteById(id);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/update")
    public Result update(User user) {
        userService.update(user);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/detail")
    public Result detail(@RequestParam Integer id) {
        User user = userService.findById(id);
        return ResultGenerator.genSuccessResult(user);
    }

    @PostMapping("/list")
    @CountTime
    public Result list(@RequestParam(defaultValue = "0") Integer page, @RequestParam(defaultValue = "0") Integer size) {
        PageHelper.startPage(page, size);
        List<User> list = userService.findAll();
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(pageInfo);
    }
    @PostMapping("/listExcel")
    @ResponseBody
    public Result listExcel(HttpServletRequest request, HttpServletResponse response){
        final String userAgent = request.getHeader("USER-AGENT");
        //response.setContentType("application/binary;charset=UTF-8");

        try {
            ServletOutputStream out = response.getOutputStream();
            try {
                String excelName = "用户信息-"+ SystemUtils.getCurrentTime("yyyyMMddhhmmss")+".xlsx";
                logger.info(excelName);
                if (userAgent.contains("Firefox")) {
                    response.setHeader("Content-Disposition", "attachment;fileName="+new String(excelName.getBytes("UTF-8"),"iso-8859-1"));
                    response.setContentType("application/vnd.ms-excel");
                }else {
                    response.setHeader("Content-Disposition", "attachment;fileName="+ URLEncoder.encode(excelName, "UTF-8"));
                    response.setContentType("application/vnd.ms-excel");

                }
            } catch (UnsupportedEncodingException e) {
                e.printStackTrace();
            }
           userService.getExcelAll(out);
            return  ResultGenerator.genSuccessResult();
        } catch (IOException e) {
            e.printStackTrace();
            return ResultGenerator.genFailResult("导出信息失败");
        }

    }
    @PostMapping("/outExcel")
    @ResponseBody
    public Result outExcel(HttpServletRequest request ,HttpServletResponse response, @RequestParam("fileSetExcel") MultipartFile file){
        //输入流
        InputStream in = null;
        //文件转换后的list数组
        ArrayList<ArrayList<Object>> listOb = null;
        if (file.isEmpty()){
            return  ResultGenerator.genFailResult("传入文件为空");
        }
        String filename = file.getOriginalFilename();
        String fileType = filename.substring(filename.lastIndexOf(".") + 1);
        if (!"xlsx".equals(fileType) && !"xls".equals(fileType)) {
            // 文件类型错误
            return ResultGenerator.genFailResult("传入文件格式不正确");
        }

        try {
            in = file.getInputStream();
            logger.info("******in*********"+in);
            // 转换poi Excel 表格对象
            if("xls".equals(fileType)){
                XSSFWorkbook wb = new XSSFWorkbook(in);//2003+
                in.close();
                try {
                    if (!userService.isExcelMistakeDataType(wb)) {
                        return ResultGenerator.genFailResult("传入的文件非模板文件");
                    }
                    listOb= userService.getBankListByExcel(wb);
                    logger.info("--------------------"+listOb);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }else if ("xlsx".equals(fileType)){
                XSSFWorkbook wb = new XSSFWorkbook(in); // 2007+
                in.close();
                try {
                    if (!userService.isExcelMistakeDataType(wb)) {
                        return ResultGenerator.genFailResult("传入的文件非模板文件");
                    }
                    logger.info("***********"+wb);
                    XSSFSheet sheetAt = wb.getSheetAt(0);

                    logger.info("i="+sheetAt);
                    listOb= userService.getBankListByExcel(wb);
                    logger.info("****************"+listOb);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }else{
                return  ResultGenerator.genFailResult("上传文件格式不正确，请提交正确的文件格式（*.xls, *.xlsx）");
            }


        } catch (IOException e) {
            return ResultGenerator.genFailResult( "文件校验失败，请上传格式正确的文件");
        }
        logger.info(listOb);
        return ResultGenerator.genSuccessResult(userService.setExcelData(listOb));
    }



}
