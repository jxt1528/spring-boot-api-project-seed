package com.company.project.dao;

import com.company.project.core.Mapper;
import com.company.project.model.User;

import java.util.HashMap;
import java.util.List;


public interface UserMapper extends Mapper<User> {
    List<User> queryAll();

    void insertOne(HashMap<String, Object> oneData);
}