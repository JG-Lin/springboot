package com.demo.springboot.service;

import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * FileName: ExcelUtil
 * Author:   LK
 * Date:     2018/7/28 11:24
 * Description: 1.0.0
 */
public abstract class Excel3To4Util {
    protected static final String xls = ".xls";
    protected static final String xlsx = ".xlsx";
    public static final String[] subTitle = {"变电站名称", "电压等级", "变电站编号", "标识器编号","经度","纬度", "产权性质(0公用、1 专用)", "采集时间"};
    public static final String[] lineTitle = {"线路名称", "线路编号", "主线名称", "主线路编号", "电压等级", "所属变电站", "结束变电站", "线路偏移", "线路分组", "维护部门", "设备型号", "产权性质", "生产厂家", "出厂编号", "投运日期", "出厂日期", "图上长度(m)", "实测长度(m)", "敷设方式", "备注", "序号", "线路类型(0:电缆工程线路、1:工程施工的线路段)", "导线类型"};
    public static final String[] lineNodeTitle = {"线路名称", "线路编号", "标识器编号","经度","纬度", "连接顺序", "上一节点距离", "上一节点敷设方式", "深度(距电缆沟、槽、顶管底深度)", "设备描述", "备注"};
    public static final String[] nodeTitle = {"标识器编号", "标识器类型(标识球、 标识钉、杆塔、 路径点)", "敷设方式(架空、电缆沟、电缆槽、埋管、顶管)", "经度", "纬度", "地面高程", "安装位置", "地理位置", "距底深度(电缆沟、槽、顶管)[cm]", "电缆沟/槽宽度[cm]", "设备描述", "同沟电缆回路数[回]", "同沟电缆名称及相序位置", "安装图", "背景图", "放置日期", "备注", "坐标编号(用于匹配外业采集的坐标)"};
    public static final int A = 0;
    public static final int B = 1;
    public static final int C = 2;
    public static final int D = 3;
    public static final int E = 4;
    public static final int F = 5;
    public static final int G = 6;
    public static final int H = 7;
    public static final int I = 8;
    public static final int J = 9;
    public static final int K = 10;
    public static final int L = 11;
    public static final int M = 12;
    public static final int N = 13;
    public static final int O = 14;
    public static final int P = 15;
    public static final int Q = 16;
    public static final int R = 17;
    public static final int S = 18;
    public static final int T = 19;
    public static final int U = 20;
    public static final int V = 21;
    public static final int W = 22;
    public static final int X = 23;
    public static final int Y = 24;
    public static final int Z = 25;
    public static final int AA = 26;
    public static final int AB = 27;
    public static final int AC = 28;
    public static final int AD = 29;
    public static final int AE = 30;

    /**
     * 解析excel表格数据
     *
     * @param cell cell
     * @return 数据转String类型
     */
    protected String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }
        //把数字当成String来读，避免出现1读成1.0的情况
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC && !HSSFDateUtil.isCellDateFormatted(cell)) {
            cell.setCellType(Cell.CELL_TYPE_STRING);
        }
        //判断数据的类型
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC: //数字
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    cellValue = DateFormatUtils.format(date, "yyyy-MM-dd HH:mm:ss");
                } else {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_STRING: //字符串  
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN: //Boolean  
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA: //公式  
                cellValue = String.valueOf(cell.getCellFormula());
                break;
            case Cell.CELL_TYPE_BLANK: //空值   
                cellValue = "";
                break;
            case Cell.CELL_TYPE_ERROR: //故障  
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
        }
        return cellValue;
    }

    /**
     * 解析excel表格数据
     *
     * @param cell cell
     * @return 数据转String类型
     */
    protected String getDateCell(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }
        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            cellValue = getStringDateValue(cell);
        } else if (cell.getCellStyle().getDataFormat() == 31) {
            Date date = cell.getDateCellValue();
            cellValue = DateFormatUtils.format(date, "yyyy-MM-dd HH:mm:ss");
        } else {
            cellValue = "";
        }
        //判断数据的类型
        return cellValue;
    }

    private String getStringDateValue(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }
        SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        SimpleDateFormat format = new SimpleDateFormat("yyyy年MM月dd日  HH:mm:ss");
        try {
            cellValue = fmt.format(format.parse(cell.getStringCellValue()));
            return cellValue;
        } catch (ParseException e) {
        }
        format = new SimpleDateFormat("yyyy年MM月dd日");
        try {
            cellValue = fmt.format(format.parse(cell.getStringCellValue()));
            return cellValue;
        } catch (ParseException e) {
        }
        format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        try {
            cellValue = fmt.format(format.parse(cell.getStringCellValue()));
            return cellValue;
        } catch (ParseException e) {
        }
        format = new SimpleDateFormat("yyyy-MM-dd");
        try {
            cellValue = fmt.format(format.parse(cell.getStringCellValue()));
            return cellValue;
        } catch (ParseException e) {
        }

        return "";
    }

    /**
     * 获取Excel的Workbook对象
     *
     * @param file 文件
     * @return Workbook
     * @throws IOException IOException
     */
    protected Workbook getExcelWordbook(File file) {
        String originalFileName = file.getPath();
        Workbook workbook = null;
        try {
            FileInputStream stream = new FileInputStream(file);
            if (originalFileName.endsWith(xls)) {
                workbook = new HSSFWorkbook(stream);
            }
            if (originalFileName.endsWith(xlsx)) {
                workbook = new XSSFWorkbook(stream);
            }
        } catch (Exception e) {
            System.out.println("文件[" + originalFileName + "]读取异常");
            e.printStackTrace();
        }
        return workbook;
    }
}
