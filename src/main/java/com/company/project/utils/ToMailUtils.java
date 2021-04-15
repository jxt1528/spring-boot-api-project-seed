package com.company.project.utils;

import com.company.project.model.Accessory;
import com.google.common.base.Joiner;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.*;
import javax.mail.util.ByteArrayDataSource;
import java.io.ByteArrayInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class ToMailUtils {
    private static final Logger logger = LoggerFactory.getLogger(ToMailUtils.class);
    private static final String SMTPHost = "smtp.163.com";

    private static final String mailFrom = "jin1282914923@163.com";

    private static final String SMTPUsername = "jin1282914923";

    private static final String SMTPPassword = "OZPXFLAEGZSWFXJL";


    public static void sendMail(String subject, List<String> content, List<String> mailTo, List<Accessory> accessoryList) {
        try {

            System.getProperties().setProperty("mail.mime.splitlongparameters", "false");

            Properties properties = new Properties();
            properties.put("mail.transport.protocol", "smtp");// 连接协议
            properties.put("mail.smtp.host", SMTPHost);// 主机名
            properties.put("mail.smtp.port", 587);// 端口号
            properties.put("mail.smtp.auth", "true");
            properties.put("mail.smtp.ssl.enable", "true");// 设置是否使用ssl安全连接 ---一般都使用
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
            BodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setText(contents);
            MimeMultipart mime = new MimeMultipart("mixed");
            //附件部分
            if (accessoryList.size() > 0) {
                MimeBodyPart contentPart = new MimeBodyPart();
                MimeMultipart contentMultipart = new MimeMultipart("related");
                for (Accessory accessory :
                        accessoryList) {
                    ByteArrayInputStream inputstream = accessory.getByteArrayInputStream();
                    String affixName = accessory.getAccessoryName();
                    if (inputstream != null) {
                        MimeBodyPart excelBodyPart = new MimeBodyPart();
                        DataSource dataSource = new ByteArrayDataSource(inputstream, "application/excel");
                        DataHandler dataHandler = new DataHandler(dataSource);
                        excelBodyPart.setDataHandler(dataHandler);
                        excelBodyPart.setFileName(MimeUtility.encodeText(affixName).replace("\\r", "").replace("\\n", ""));
                        contentMultipart.addBodyPart(excelBodyPart);
                    }
                }
                contentPart.setContent(contentMultipart);
                mime.addBodyPart(contentPart);
            }
            mime.addBodyPart(messageBodyPart);
            message.setContent(mime);

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
}
