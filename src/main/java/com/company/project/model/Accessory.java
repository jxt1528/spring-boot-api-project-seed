package com.company.project.model;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

public class Accessory {
    //附件名称 （包含后缀）
    private String accessoryName;
    //附件 输入流
    private ByteArrayInputStream byteArrayInputStream;

    /**
     * 将list 转换excel
     *
     * @param dataList      数据集合
     * @param dataType      数据类型
     * @param sheetName     页名称
     * @param accessoryName 附件名称 （包含后缀）
     * @throws IOException
     */
    public Accessory(List<?> dataList, Class<?> dataType, String sheetName, String accessoryName) throws IOException {
        this.accessoryName = accessoryName;
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        EasyExcel.write(out, dataType).sheet(sheetName).doWrite(dataList);
        this.byteArrayInputStream = new ByteArrayInputStream(out.toByteArray());
        out.close();
    }


    public Accessory(List<?> dataList, Class<?> dataType, String sheetName, String accessoryName,
                     List<?> dataList2, Class<?> dataType2, String sheetName2, String accessoryName2) throws IOException {
        this.accessoryName = accessoryName;
        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        EasyExcel.write(out, dataType).sheet(sheetName).doWrite(dataList);
//        this.byteArrayInputStream = new ByteArrayInputStream(out.toByteArray());
        ExcelWriter excelWriter = EasyExcel.write(out).build();
        WriteSheet writeSheet   = EasyExcel.writerSheet(0, sheetName).head(dataType).build();
        excelWriter.write(dataList, writeSheet);

        WriteSheet writeSheet2   = EasyExcel.writerSheet(1, sheetName2).head(dataType2).build();
        excelWriter.write(dataList2, writeSheet2);
        excelWriter.finish();
        this.byteArrayInputStream = new ByteArrayInputStream(out.toByteArray());
        out.close();
    }

    public Accessory() {
    }


    public String getAccessoryName() {
        return accessoryName;
    }

    public void setAccessoryName(String accessoryName) {
        this.accessoryName = accessoryName;
    }

    public ByteArrayInputStream getByteArrayInputStream() {
        return byteArrayInputStream;
    }

    public void setByteArrayInputStream(ByteArrayInputStream byteArrayInputStream) {
        this.byteArrayInputStream = byteArrayInputStream;
    }

}
