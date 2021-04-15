package com.company.project.service.impl;

import cn.hutool.Hutool;
import com.company.project.dao.UserMapper;
import com.company.project.model.User;
import com.company.project.service.TaskTimeService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;

import java.time.LocalDateTime;
import java.util.List;
@Service
//@EnableScheduling
public class TaskTimeServiceImpl implements TaskTimeService {

    @Autowired
    UserMapper userMapper;
    @Override
    @Scheduled(cron = "0/5 * * * * ? ")
    public List<User> selectAll() {
        List<User> userList = userMapper.selectAll();
        System.out.println(LocalDateTime.now()+"执行了");
        return userList;
    }
}
