package com.company.project.utils;


import com.alibaba.excel.EasyExcel;
import com.google.common.base.Joiner;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang.time.DateFormatUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.*;
import javax.mail.util.ByteArrayDataSource;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Properties;

/**
 * 邮件发送工具类
 * @author songjunhao
 */
public class MailUtils {

    private static final Logger logger = LoggerFactory.getLogger(MailUtils.class);

    private static final String SMTPHost = "10.111.10.11";

    private static final String mailFrom = "XSZC_Sys@picclife.cn";

    private static final String SMTPUsername = "XSZC_Sys";

    private static final String SMTPPassword = "xszc@1234";


    //生成excel
    public static <T> ByteArrayInputStream generateExcel(List<T> peopleManangerAlreadyInPeopleInfos, Class<T> clazz, String sheetName) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        EasyExcel.write(out, clazz).sheet(sheetName).doWrite(peopleManangerAlreadyInPeopleInfos);
        ByteArrayInputStream excelStream = new ByteArrayInputStream(out.toByteArray());
        out.close();
        return excelStream;
    }

    /**
     * 发送Excel邮件，自动把list 转为Excel 发送
     *
     * @param contentList 数据
     * @param clazz       数据类型
     * @param subject     主题
     * @param mailTo      收件人
     * @param <T>         泛型
     */
    public static <T> void sendExcelMail(List<T> contentList, Class<T> clazz, String subject, List<String> mailTo) {
        ByteArrayInputStream byteArrayInputStream = null;

        ArrayList<String> context = new ArrayList<>();
        if (CollectionUtils.isNotEmpty(contentList)) {
            try {
                byteArrayInputStream = generateExcel(contentList, clazz, subject);
            } catch (IOException e) {
                logger.error("生成Excel异常");
            }
            context.add("包含数据：" + contentList.size() + "条");
        } else {
            context.add("无数据");
        }

        if (byteArrayInputStream != null) {
            MailUtils.sendMail(subject, context, mailTo, byteArrayInputStream, subject + ".xlsx");
        } else {
            MailUtils.sendMail(subject + "失败", context, mailTo);
        }
    }

    /**
     * @param subject     标题
     * @param content     邮件内容
     * @param mailTo      收件人
     * @param inputstream 附件
     * @param affixName   附件名
     */
    public static void sendMail(String subject, List<String> content, List<String> mailTo, ByteArrayInputStream inputstream, String affixName) {
        try {

            System.getProperties().setProperty("mail.mime.splitlongparameters", "false");

            Properties properties = new Properties();
            properties.put("mail.transport.protocol", "smtp");// 连接协议
            properties.put("mail.smtp.host", SMTPHost);// 主机名
            properties.put("mail.smtp.port", 587);// 端口号
            properties.put("mail.smtp.auth", "true");
            properties.put("mail.smtp.ssl.enable", "false");// 设置是否使用ssl安全连接 ---一般都使用
            properties.put("mail.debug", "true");// 设置是否显示debug信息 true 会在控制台显示相关信息

            // 得到回话对象
            Session session = Session.getInstance(properties);
            // 获取邮件对象
            Message message = new MimeMessage(session);
            // 设置发件人邮箱地址
            message.setFrom(new InternetAddress(mailFrom));

            ArrayList<InternetAddress> internetAddresslist = new ArrayList<>();

            for (String s : mailTo) {
                internetAddresslist.add(new InternetAddress(s));
            }

            InternetAddress[] mailGroup = new InternetAddress[mailTo.size()];
            // 设置收件人邮箱地址
            message.setRecipients(Message.RecipientType.TO, internetAddresslist.toArray(mailGroup));

            // 设置邮件标题
            message.setSubject(subject);
            // 设置邮件内容
            String contents = Joiner.on("\r\n").join(content);
            message.setText(contents);
            //附件部分
            if (inputstream != null) {
                MimeBodyPart contentPart = new MimeBodyPart();
                MimeMultipart contentMultipart = new MimeMultipart("related");
                MimeBodyPart excelBodyPart = new MimeBodyPart();
                DataSource dataSource = new ByteArrayDataSource(inputstream, "application/excel");
                DataHandler dataHandler = new DataHandler(dataSource);
                excelBodyPart.setDataHandler(dataHandler);
                excelBodyPart.setFileName(MimeUtility.encodeText(affixName).replace("\\r","").replace("\\n",""));
                contentMultipart.addBodyPart(excelBodyPart);
                contentPart.setContent(contentMultipart);
                MimeMultipart mime = new MimeMultipart("mixed");
                mime.addBodyPart(contentPart);
                message.setContent(mime);
            }

            //保存邮件
            message.saveChanges();
            // 得到邮差对象
            Transport transport = session.getTransport();
            // 连接自己的邮箱账户
            transport.connect(SMTPUsername, SMTPPassword);// 密码为QQ邮箱开通的stmp服务后得到的客户端授权码
            // 发送邮件
            transport.sendMessage(message, message.getAllRecipients());
            transport.close();

        } catch (Exception e) {
            logger.error("发送邮件异常", e);
        }

    }

    public static void sendMail(String subject, List<String> content, List<String> mailTo) {
        try {

            Properties properties = new Properties();
            properties.put("mail.transport.protocol", "smtp");// 连接协议
            properties.put("mail.smtp.host", SMTPHost);// 主机名
            properties.put("mail.smtp.port", 587);// 端口号
            properties.put("mail.smtp.auth", "true");
            properties.put("mail.smtp.ssl.enable", "false");// 设置是否使用ssl安全连接 ---一般都使用
            properties.put("mail.debug", "true");// 设置是否显示debug信息 true 会在控制台显示相关信息

            // 得到回话对象
            Session session = Session.getInstance(properties);
            // 获取邮件对象
            Message message = new MimeMessage(session);
            // 设置发件人邮箱地址
            message.setFrom(new InternetAddress(mailFrom));

            ArrayList<InternetAddress> internetAddresslist = new ArrayList<>();

            for (String s : mailTo) {
                internetAddresslist.add(new InternetAddress(s));
            }

            InternetAddress[] mailGroup = new InternetAddress[mailTo.size()];
            // 设置收件人邮箱地址
            message.setRecipients(Message.RecipientType.TO, internetAddresslist.toArray(mailGroup));

            // 设置邮件标题
            message.setSubject(subject);
            // 设置邮件内容
            String contents = Joiner.on("\r\n").join(content);
            message.setText(contents);

            // 得到邮差对象
            Transport transport = session.getTransport();
            // 连接自己的邮箱账户
            transport.connect(SMTPUsername, SMTPPassword);// 密码为QQ邮箱开通的stmp服务后得到的客户端授权码
            // 发送邮件
            transport.sendMessage(message, message.getAllRecipients());
            transport.close();

        } catch (Exception e) {
            logger.error("发送邮件异常", e);
        }

    }

    public static void sendHtmlMail(String subject,List<String> headerList, List<List<String>> content, List<String> mailTo) {
        try {

            Properties properties = new Properties();
            properties.put("mail.transport.protocol", "smtp");// 连接协议
            properties.put("mail.smtp.host", SMTPHost);// 主机名
            properties.put("mail.smtp.port", 587);// 端口号
            properties.put("mail.smtp.auth", "true");
            properties.put("mail.smtp.ssl.enable", "false");// 设置是否使用ssl安全连接 ---一般都使用
            properties.put("mail.debug", "true");// 设置是否显示debug信息 true 会在控制台显示相关信息

            // 得到回话对象
            Session session = Session.getInstance(properties);
            // 获取邮件对象
            Message message = new MimeMessage(session);
            // 设置发件人邮箱地址
            message.setFrom(new InternetAddress(mailFrom));

            ArrayList<InternetAddress> internetAddresslist = new ArrayList<>();

            for (String s : mailTo) {
                internetAddresslist.add(new InternetAddress(s));
            }

            InternetAddress[] mailGroup = new InternetAddress[mailTo.size()];
            // 设置收件人邮箱地址
            message.setRecipients(Message.RecipientType.TO, internetAddresslist.toArray(mailGroup));

            // 设置邮件标题
            message.setSubject(subject);

            // 设置邮件内容
            MimeBodyPart htmlPart = new MimeBodyPart();
            htmlPart.setContent(buildContent(subject, headerList, content),"text/html; charset=utf-8");
            // MiniMultipart类是一个容器类，包含MimeBodyPart类型的对象
            Multipart mp = new MimeMultipart("alternative");//mixed related alternative
            mp.addBodyPart(htmlPart);

            message.setContent(mp);

            // 得到邮差对象
            Transport transport = session.getTransport();
            // 连接自己的邮箱账户
            transport.connect(SMTPUsername, SMTPPassword);// 密码为QQ邮箱开通的stmp服务后得到的客户端授权码
            // 发送邮件
            transport.sendMessage(message, message.getAllRecipients());
            transport.close();

        } catch (Exception e) {
            logger.error("发送邮件异常", e);
        }

    }


    private static String buildContent(String subject, List<String> headerList, List<List<String>> contentList) {
        //加载邮件html模板
        String buffer = "<body style=\"color: #666; font-size: 14px; font-family: 'Open Sans',Helvetica,Arial,sans-serif;\">\n" +
                "<div class=\"box-content\" style=\"width: 80%; margin: 20px auto; max-width: 800px; min-width: 600px;\">\n" +
                "    <div class=\"header-tip\" style=\"font-size: 12px;\n" +
                "                                   color: #aaa;\n" +
                "                                   text-align: right;\n" +
                "                                   padding-right: 25px;\n" +
                "                                   padding-bottom: 10px;\">\n" +
                "        Confidential - Scale Alarm Use Only\n" +
                "    </div>\n" +
                "    <div class=\"info-top\" style=\"padding: 15px 25px;\n" +
                "                                 border-top-left-radius: 10px;\n" +
                "                                 border-top-right-radius: 10px;\n" +
                "                                 background: {0};\n" +
                "                                 color: #fff;\n" +
                "                                 overflow: hidden;\n" +
                "                                 line-height: 32px;\">\n" +
                "    <div style=\"color:#010e07\"><strong>" + subject + "</strong></div>\n" +
                "    </div>\n" +
                "    <div class=\"info-wrap\" style=\"border-bottom-left-radius: 10px;\n" +
                "                                  border-bottom-right-radius: 10px;\n" +
                "                                  border:1px solid #ddd;\n" +
                "                                  overflow: hidden;\n" +
                "                                  padding: 15px 15px 20px;\">\n" +
                "        <div class=\"tips\" style=\"padding:15px;\">\n" +
                "            <p style=\" list-style: 160%; margin: 10px 0;\">各位好,</p>\n" +
                "            <p style=\" list-style: 160%; margin: 10px 0;\">{1}</p>\n" +
                "        </div>\n" +
                "        <div class=\"time\" style=\"text-align: right; color: #999; padding: 0 15px 15px;\">{2}</div>\n" +
                "        <br>\n" +
                "        <table class=\"list\" style=\"width: 100%; border-collapse: collapse; border-top:1px solid #eee; font-size:12px;\">\n" +
                "            <thead>\n" +
                "            <tr style=\" background: #fafafa; color: #333; border-bottom: 1px solid #eee;\">\n" +
                "                {3}\n" +
                "            </tr>\n" +
                "            </thead>\n" +
                "            <tbody>\n" +
                "            {4}\n" +
                "            </tbody>\n" +
                "        </table>\n" +
                "    </div>\n" +
                "</div>\n" +
                "</body>";


        String contentText = "以下是 " + subject + " , 敬请查看.";
        //邮件表格header
        String header = "";
        StringBuilder stringBuilder = new StringBuilder();
        for (String headerEntity : headerList) {
            stringBuilder.append("<td>").append(headerEntity).append("</td>");
        }
        header = stringBuilder.toString();

        StringBuilder linesBuffer = new StringBuilder();
        StringBuilder lineInfo = new StringBuilder();
        for (List<String> contentListEntity : contentList ) {
            lineInfo.append("<tr>");
            for (String contentEntity : contentListEntity) {
                lineInfo.append("<td>").append(contentEntity).append("</td>");
            }
            lineInfo.append("</tr>");
            linesBuffer.append(lineInfo.toString());
            lineInfo.setLength(0);
        }
//        linesBuffer.append("<tr><td>" + "myNamespace" + "</td><td>" + "myServiceName" + "</td><td>" + "myscaleResult" + "</td>" +
//                "<td>" + "mReason" + "</td><td>" + "my4" + "</td></tr>");

        //绿色
        String emailHeadColor = "#10fa81";
        String date = DateFormatUtils.format(new Date(), "yyyy/MM/dd HH:mm:ss");
        //填充html模板中的五个参数
        String htmlText = MessageFormat.format(buffer, emailHeadColor, contentText, date, header, linesBuffer.toString());

        //改变表格样式
        htmlText = htmlText.replaceAll("<td>", "<td style=\"padding:6px 10px; line-height: 150%;\">");
        htmlText = htmlText.replaceAll("<tr>", "<tr style=\"border-bottom: 1px solid #eee; color:#666;\">");
        return htmlText;
    }


}
