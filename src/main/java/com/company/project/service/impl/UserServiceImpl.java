package com.company.project.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.company.project.core.Mapper;
import com.company.project.dao.UserMapper;
import com.company.project.model.User;
import com.company.project.service.UserService;
import com.company.project.core.AbstractService;
import com.company.project.utils.ImportExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import javax.servlet.ServletOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * Created by CodeGenerator on 2021/03/10.
 */
@Service
@Transactional
public class UserServiceImpl extends AbstractService<User> implements UserService {
    @Resource
    private UserMapper userMapper;
    @Autowired
    private ImportExcelUtils importexcelutils;
    private static final Logger logger = LoggerFactory.getLogger(UserServiceImpl.class);
    private final String[] titles = { "用户名", "密码", "姓名", "性别", "注册日期"};

    @Override
    public void getExcelAll(ServletOutputStream out) {
    List<User> userList = new ArrayList<>();
        try {
          userList = userMapper.queryAll();
        } catch (Exception e) {
            e.printStackTrace();
        }
        ExcelWriter write = EasyExcel.write(out, User.class).build();
        WriteSheet sheet = EasyExcel.writerSheet("用户信息").build();
        write.write(userList, sheet);
        write.finish();
    }

    @Override
    public boolean isExcelMistakeDataType(HSSFWorkbook wb) throws Exception {
        return importexcelutils.isExcelMistakeDataType(wb, titles);
    }

    @Override
    public boolean isExcelMistakeDataType(XSSFWorkbook wb) throws Exception {
        return importexcelutils.isExcelMistakeDataType(wb, titles);
    }

    @Override
    public ArrayList<ArrayList<Object>> getBankListByExcel(HSSFWorkbook work) throws Exception {
        return importexcelutils.getBankListByExcel(work);
    }

    @Override
    public ArrayList<ArrayList<Object>> getBankListByExcel(XSSFWorkbook work) throws Exception {
        return importexcelutils.getBankListByExcel(work);
    }

    @Override
    public Map<String, Object> setExcelData(ArrayList<ArrayList<Object>> data) {
        Map<String, Object> map = new HashMap<>();
        ArrayList<HashMap<String, Object>> recordData = new ArrayList<HashMap<String, Object>>();
        int i = 0;
        try {
            for (ArrayList<Object> arrayList : data) {
                i++;
                if (arrayList.size() <= 1) {
                    logger.info("第" + i + "条数据为空，不导入");

                    continue;
                }
                if (arrayList.get(1) == null && "".equals(arrayList.get(1))) {
                    logger.info("第" + i + "条数据无macId，导入失败");
                    continue;
                }

                HashMap<String, Object> oneData = new HashMap<String, Object>();
                oneData.put("username", arrayList.get(0));
                oneData.put("password", arrayList.get(1));
                oneData.put("nickName", arrayList.get(2));
                oneData.put("sex", arrayList.get(3));
                oneData.put("registerDate", arrayList.get(4));

                // 添加设备基础信息
                recordData.add(oneData);
                userMapper.insertOne(oneData);
            }


        } catch (Exception e) {
            // TODO: handle exception
            logger.debug("导入失败，下标>>>>" + i);
            return null;
        }

        return map;
    }
}
