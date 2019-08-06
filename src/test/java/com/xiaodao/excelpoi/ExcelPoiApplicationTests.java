package com.xiaodao.excelpoi;


import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.regex.Matcher;
import java.util.regex.Pattern;


@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelPoiApplicationTests {

    @Test
    public void contextLoads() {
    }

    private Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}",
                Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }
    @Test
    public  void testMatcher() {

        String a= "${xiaodao}${gaolei}";
        Matcher matcher = this.matcher(a);

        while (matcher.find()) {
            System.out.println(matcher.group(1));
        }
    }




}
