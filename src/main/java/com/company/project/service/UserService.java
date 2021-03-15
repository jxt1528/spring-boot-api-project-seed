package com.company.project.service;
import com.company.project.model.User;
import com.company.project.core.Service;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import java.util.ArrayList;
import java.util.Map;


/**
 * Created by CodeGenerator on 2021/03/10.
 */
public interface UserService extends Service<User> {
    public  void getExcelAll(ServletOutputStream out); //导出excel
    public boolean isExcelMistakeDataType(HSSFWorkbook wb) throws Exception;
    public boolean isExcelMistakeDataType( XSSFWorkbook wb) throws Exception;
    public ArrayList<ArrayList<Object>> getBankListByExcel(HSSFWorkbook work) throws Exception;
    public ArrayList<ArrayList<Object>> getBankListByExcel(XSSFWorkbook work) throws Exception;
    public Map<String,Object> setExcelData(ArrayList<ArrayList<Object>> data);
}
