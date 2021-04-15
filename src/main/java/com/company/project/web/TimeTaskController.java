package com.company.project.web;

import cn.hutool.Hutool;
import com.company.project.core.Result;
import com.company.project.core.ResultGenerator;
import com.company.project.model.Accessory;
import com.company.project.model.User;
import com.company.project.service.TaskTimeService;
import com.company.project.utils.SystemUtils;
import com.company.project.utils.ToMailUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;


@RestController
@RequestMapping("/task")

public class TimeTaskController {
    private final Log logger = LogFactory.getLog(TimeTaskController.class);
    @Autowired
    TaskTimeService taskTimeService;
    //定时任务测试 每隔10查询一下用户数据
    @RequestMapping("/test2")
    public Result taskTest() throws IOException {
        List<User> userList = taskTimeService.selectAll();
        //获取发送邮件的日期
        Date date = new Date();
        String[] weekDays = {"星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"};
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        String week = weekDays[calendar.get(Calendar.DAY_OF_WEEK) - 1];
        //将date类型转传承yyyy-MM-dd字符串
        String countDate = SystemUtils.getCountDate(date, 0, 0, 0);
        LocalDate localDate = LocalDate.parse(countDate);
        String datelast=  localDate.getMonthValue()+"月"+localDate.getDayOfMonth()+"日";
        //设置发送邮件的额标题
        String title = datelast+"("+week+")用户表信息";
        //设置发送邮件的内容
        List<String> content = new ArrayList<>();
        content.add("亲爱的领导");
        content.add("       截止"+datelast+"我们已经将所有用户信息表统计完毕");
        //设置将要发送的邮箱
        List<String> toTail = new ArrayList<>();
        toTail.add("1282914923@qq.com");
        //设置需要发送的附件
        List<Accessory> accessoryList = new ArrayList<>();
        Accessory accessory =new Accessory(userList,User.class,"用户信息表","用户表.xlsx");
        accessoryList.add(accessory);
        ToMailUtils.sendMail(title, content, toTail, accessoryList);


        return ResultGenerator.genSuccessResult();
    }

    @RequestMapping("test1")
    public Result test2(){
        List<User> userList = taskTimeService.selectAll();
        return ResultGenerator.genSuccessResult(userList);
    }




    public static void main(String[] args) {


    }

}
