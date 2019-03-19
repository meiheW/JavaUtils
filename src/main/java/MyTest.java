import myutils.DateUtil;
import myutils.ExcelUtil;
import myutils.StringUtil;

public class MyTest {

    public static void main(String[] args){

        System.out.println(StringUtil.timeString());
        System.out.println(DateUtil.now());
        ExcelUtil.downloadDemo(new String[]{"pageName", "h5Url", "wechatUrl", "alipayUrl", "baiduUrl"}, null, "GeneralQRCodeExample.xls");


    }


}
