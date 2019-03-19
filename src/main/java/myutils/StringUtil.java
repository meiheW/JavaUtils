package myutils;

import java.text.SimpleDateFormat;
import java.util.Date;

public class StringUtil {


    /**
     * 判断指定字符串是否为空
     * @param str 待检查字符串
     * @return T/F
     */
    public static boolean isEmpty(String str){
        if(str==null || str.length()==0)
            return true;

        for (int i = 0; i < str.length(); i++) {
            if ((Character.isWhitespace(str.charAt(i)) == false)) {
                return false;
            }
        }

        return true;
    }

    /**
     * 判断对象是否为数字型字符串，包含负数
     * @param obj 待检查对象
     * @return T/F
     */
    public static boolean isNumeric(Object obj) {
        if (obj == null) {
            return false;
        }
        char[] chars = obj.toString().toCharArray();
        int length = chars.length;
        if(length < 1)
            return false;

        int i = 0;
        if(length > 1 && chars[0] == '-')
            i = 1;

        for (; i < length; i++) {
            if (!Character.isDigit(chars[i])) {
                return false;
            }
        }
        return true;
    }

    /**
     * 按照yyyyMMddHH格式返回日期
     * @return
     */
    public static String timeString(){
        return timeString("yyyyMMddHH");
    }

    /**
     * 格式化输出日期
     * @param format
     * @return
     */
    public static String timeString(String format){
        SimpleDateFormat sfDate = new SimpleDateFormat(format);
        return sfDate.format(new Date());
    }



}
