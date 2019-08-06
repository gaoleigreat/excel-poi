package com.xiaodao.excelpoi.util;

import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelUtil {

    /*正则匹配${}*/
    public static Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}",
                Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }


    /**
     * 正则匹配字符串是否含有
     *
     * @param str
     * @return
     */
    public static boolean isMatcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}",
                Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher.find();
    }

    /**
     * 获取文件输入流，输入可以使链接，也可以是路径
     *
     * @param filePath
     * @return
     * @throws Exception
     */
    public static InputStream getFileStream(String filePath) throws Exception {
        if (StringUtils.isNotBlank(filePath) && filePath.startsWith("http")) {
            URL url = new URL(filePath);
            //打开链接
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            //设置请求方式为"GET"
            conn.setRequestMethod("GET");
            //超时响应时间为5秒
            conn.setConnectTimeout(5 * 1000);
            //通过输入流获取图片数据
            InputStream is = conn.getInputStream();
            return is;
        } else if (StringUtils.isNotBlank(filePath)) {
            InputStream is = new FileInputStream(filePath);
            return is;
        } else {
            return null;
        }
    }

    /**
     * 关闭输入流
     * @param is
     */
    public static void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流
     *
     * @param os
     */
    public static void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }




}
