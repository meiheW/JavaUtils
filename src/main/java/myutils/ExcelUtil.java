package myutils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {
    private final static String EXC_2003L = ".xls";    //2003- 版本的excel
    private final static String EXC_2007U = ".xlsx";   //2007+ 版本的excel

    /**
     * 检验文件格式
     */
    public static void checkFile(MultipartFile file, String fileName, String[] columns) throws Exception {
        InputStream in = file.getInputStream();
        //创建Excel工作薄
        Workbook work = getWorkbook(in, fileName);
        if (null == work) {
            return;
        }
        //获取sheet1
        Sheet sheet = work.getSheetAt(0);
        if (sheet == null) {
            return;
        }
        //获取表头
        Row row = sheet.getRow(0);
        for (int y = 0; y < row.getLastCellNum(); y++) {
            Cell cell = row.getCell(y);
            if (!getCellValue(cell).equals(columns[y])) {
                return;
            }

        }
    }

    /**
     * 获取IO流中的数据，组装成List<List<Object>>对象
     *
     * @param file,fileName 文件， 文件名
     * @return 列表
     */
    public static List<List<Object>> getBankListByExcel(MultipartFile file, String fileName) throws Exception {

        List<List<Object>> list = null;
        InputStream in = file.getInputStream();
        //创建Excel工作薄
        Workbook work = getWorkbook(in, fileName);
        if (null == work) {
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;

        list = new ArrayList<List<Object>>();
        //遍历Excel中所有的sheet
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            sheet = work.getSheetAt(i);
            if (null == sheet) {
                continue;
            }
            //获取表头列数
            int count = sheet.getRow(0).getLastCellNum();
            //遍历当前sheet中的所有行
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                if (null == row) {
                    continue;
                }
                //遍历所有的列
                List<Object> li = new ArrayList<>();
                for (int y = 0; y < count; y++) {
                    cell = row.getCell(y);
                    if (null == cell) {
                        li.add("");
                    } else {
                        li.add(getCellValue(cell));
                    }

                }
                list.add(li);
            }
        }
        in.close();
        return list;
    }


    /**
     * 获取IO流中的数据，组装成List<List<List<Object>>>对象
     * @param file,fileName 文件， 文件名
     * @return 列表
     */
    public static List<List<List<Object>>> getBaseListByExcel3D(MultipartFile file, String fileName) throws Exception{

        List<List<List<Object>>> resList = new ArrayList<>();

        InputStream in = file.getInputStream();
        //创建Excel工作薄
        Workbook work = getWorkbook(in,fileName);
        if(null == work){
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;

        //遍历Excel中所有的sheet
        for (int i = 0; i < work.getNumberOfSheets(); i++) {

            List<List<Object>> list = new ArrayList<List<Object>>();
            sheet = work.getSheetAt(i);
            if(null == sheet){
                continue;
            }

            //遍历当前sheet中的所有行
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                if(null == row){
                    continue;
                }
                //遍历所有的列
                List<Object> li = new ArrayList<>();

                for (int y = 0; y < row.getLastCellNum(); y++) {
                    cell = row.getCell(y);
                    if(null == cell){
                        li.add("");
                    }else {
                        li.add(getCellValue(cell));
                    }

                }
                if(!StringUtil.isEmpty((String)li.get(0))) {
                    list.add(li);
                }else{
                    break;
                }
            }

            resList.add(list);
        }
        in.close();
        return resList;
    }


    /**
     * 描述：根据文件后缀，自适应上传文件的版本
     * @param inStr,fileName 文件输入流，文件名
     * @return excel文件
     * @throws Exception 异常
     */
    public static Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (EXC_2003L.equals(fileType)) {
            wb = new HSSFWorkbook(inStr);  //2003-
        } else if (EXC_2007U.equals(fileType)) {
            wb = new XSSFWorkbook(inStr);  //2007+
        } else {
            throw new Exception("解析的文件格式有误！");
        }
        return wb;
    }

    /**
     * 描述：对表格中数值进行格式化
     *
     * @param cell 单元格
     * @return 单元格的数据
     */
    public static Object getCellValue(Cell cell) {
        Object value = null;
        DecimalFormat df = new DecimalFormat("0");  //格式化number String字符
        SimpleDateFormat sdf = new SimpleDateFormat("yyy-MM-dd");  //日期格式化
        DecimalFormat df2 = new DecimalFormat("0.00");  //格式化数字

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if ("General".equals(cell.getCellStyle().getDataFormatString())) {
                    value = df.format(cell.getNumericCellValue());
                } else if ("m/d/yy".equals(cell.getCellStyle().getDataFormatString())) {
                    value = sdf.format(cell.getDateCellValue());
                } else {
                    value = df2.format(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_BLANK:
                value = "";
                break;
            default:
                break;
        }
        return value;
    }

    /**
     * 下载EXCEL文件
     * @param column 每列第一行构成的字符串数组
     * @param value EXCEL表内的数据，二维数组
     * @param fileName 文件名
     */
    public static void downloadDemo(String[] column, List<List<String>> value, String fileName) {

        HSSFWorkbook wb = getHssfworkbook(fileName, column, value);
        try {
            HttpServletResponse response = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getResponse();
            setResponseHeader(response, fileName);

            OutputStream os = response.getOutputStream();
            wb.write(os);
            os.flush();
            os.close();
        } catch (Exception e) {
            return;
        }
    }

    /**
     * 构建HSSFWorkbook操作EXCEL读写
     * @param sheetName 文件名
     * @param title 标题数组
     * @param value EXCEL表内的数据，二维数组
     * @return
     */
    public static HSSFWorkbook getHssfworkbook(String sheetName, String[] title, List<List<String>> value) {

        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();

        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(sheetName);

        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        HSSFRow row = sheet.createRow(0);

        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFillPattern(HSSFCellStyle.NO_FILL);

        //声明列对象
        HSSFCell cell = null;

        //创建标题行
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        //创建内容
        if (value != null) {
            for (int i = 0; i < value.size(); i++) {
                if (null != value.get(i)) {
                    row = sheet.createRow(i + 1);
                    for (int j = 0; j < value.get(i).size(); j++) {
                        //将内容按顺序赋给对应的列对象
                        row.createCell(j).setCellValue(value.get(i).get(j));
                    }
                }

            }
        }

        return wb;
    }

    public static void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            try {
                fileName = new String(fileName.getBytes(), "ISO8859-1");
            } catch (UnsupportedEncodingException e) {
                return;
            }
            response.setContentType("application/octet-stream;charset=ISO8859-1");
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
            response.addHeader("Pargam", "no-cache");
            response.addHeader("Cache-Control", "no-cache");
        } catch (Exception ex) {
            return;
        }
    }

    /**
     * 下载多页EXCEL文件
     * @param response
     * @throws IOException
     */
    public void DownloadSheetDemo(HttpServletResponse response) throws IOException {

        String fileName = "ProductsDemo.xls";
        HSSFWorkbook wb = getHssfworkbookConfig();

        response.setContentType("application/octet-stream;charset=ISO8859-1");
        response.setHeader("Content-Disposition", "attachment;filename="+ fileName);
        response.addHeader("Pargam", "no-cache");
        response.addHeader("Cache-Control", "no-cache");

        OutputStream os = response.getOutputStream();
        wb.write(os);
        os.flush();
        os.close();
        return ;

    }

    public static HSSFWorkbook getHssfworkbookConfig(){
        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();

        String sheetName1 = "商品基本配置";
        String sheetName2 = "商品原价配置";
        String sheetName3 = "商品折扣价配置";
        String sheetName4 = "商品底价配置";
        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet1 = wb.createSheet(sheetName1);
        HSSFSheet sheet2 = wb.createSheet(sheetName2);
        HSSFSheet sheet3 = wb.createSheet(sheetName3);
        HSSFSheet sheet4 = wb.createSheet(sheetName4);

        String[] title1 = {"商品ID（BU）", "商品名称", "商品简介", "奖品描述", "图片地址", "详情地址", "门市价", "排序编号",
                "营销标签", "所在城市", "所属主题", "参与人数初始值", "新客策略", "老客策略", "推荐商品ID（市场）", "生效时间", "失效时间"};
        String[] title2 = {"原价", "新客原价名称", "老客原价名称", "原价购买连接"};
        String[] title3 = {"新客折扣价", "老客折扣价", "新客折扣价名称", "老客折扣价名称", "新客折扣价优惠券策略", "老客折扣价优惠券策略",
                "折扣价购买连接", "折扣价总库存", "折扣价达成人数"};
        String[] title4 = {"新客底价", "老客底价", "新客底价名称", "老客底价名称", "新客底价优惠券策略", "老客底价优惠券策略",
                "底价购买连接", "底价总库存", "底价达成人数"};

        HSSFSheet[] sheetArr = new HSSFSheet[]{sheet1, sheet2, sheet3, sheet4};
        String[][] titleArr = new String[][]{title1, title2, title3, title4};

        // 第三步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        //声明列对象
        HSSFCell cell = null;

        // 第四步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制


        for(int i=0; i<sheetArr.length; i++){
            HSSFRow row = sheetArr[i].createRow(0);
            for(int j = 0;j<titleArr[i].length;j++){
                cell = row.createCell(j);
                cell.setCellValue(titleArr[i][j]);
                cell.setCellStyle(style);
            }
        }
        //创建标题行

        return wb;
    }


}
