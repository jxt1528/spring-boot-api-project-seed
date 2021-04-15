package com.company.project.utils;

import org.springframework.stereotype.Component;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

@Component
public class SystemUtils {

    /**
     * yyyyMMddhhmmss
     *
     * @获取当前时间
     */
    public static String getCurrentTime() {

        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String dataTime = df.format(new Date()).toString();
        return dataTime;
    }

    public static String getCurrentTime(String dateType) {
        try {
            SimpleDateFormat df = new SimpleDateFormat(dateType);
            String dataTime = df.format(new Date()).toString();
            dateType = dataTime;
        } catch (Exception e) {
            return dateType + "格式不正确";
        }

        return dateType;

    }
    public static String getCountDate(Date date, int YEAR, int MONTH, int DAY) {
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
        String dates = "";
        try {
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(date);
            calendar.add(Calendar.YEAR, +YEAR);
            calendar.add(Calendar.MONTH, +MONTH);
            calendar.add(Calendar.DATE, +DAY);
            dates = df.format(calendar.getTime()).toString();
        } catch (Exception e) {
            // TODO: handle exception
        }
        return dates;

    }
}
