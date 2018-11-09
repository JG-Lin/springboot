package com.demo.springboot.controller;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * FileName: DataCache
 * Author:   LK
 * Date:     2018/7/28 11:24
 * Description: 1.0.0
 */
public class Excel3To4DataCache {

//            public static String logFilePath = "D:/excel/log/";
    public static String logFilePath = "./log/";

    public static FileOutputStream out;
    /**
     * 保存变电站编号
     */
    public static List<String> subNoList = new CopyOnWriteArrayList<>();
    /**
     * 保存变电站标识器编号
     */
    public static Map<String, String> subMarkerNoMap = new ConcurrentHashMap<>();
    /**
     * 保存线路编号
     */
    public static List<String> lineNoList = new CopyOnWriteArrayList<>();
    /**
     * 保存线路主线编号
     */
    public static Map<String, String> lineParentNoList = new ConcurrentHashMap<>();

    /**
     * 创建日志文件，日志文件夹
     */
    public static void createLogFile() {
        Date date = new Date(System.currentTimeMillis());
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        File file = new File(logFilePath);
        file.mkdirs();
        File logFile = new File(logFilePath + simpleDateFormat.format(date) + "log.txt");
        try {
            if (!logFile.exists()) {
                logFile.createNewFile();
            }
            out = new FileOutputStream(logFile, true);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 根据所传exception写入日志文件
     *
     * @param e
     */
    public synchronized static void logError(Exception e) {
        try {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);
            out.write(sw.toString().getBytes());
            out.flush();
        } catch (IOException e2) {
            e2.printStackTrace();
        }
    }

    /**
     * 销毁文件流
     *
     * @throws IOException
     */
    public static void destroy() throws IOException {
        out.close();
    }

    /**
     * 判断变电站编号是否已经存在
     *
     * @param subNo 变电站编号
     * @return 已经存在对应的变电站编号返回true
     */
    public static Boolean isSubNo(String subNo) {
        return subNoList.contains(subNo);
    }


    /**
     * 判断线路编号是否存在
     *
     * @param lineNo 线路编号
     * @return 已经存在对应的线路编号返回true
     */
    public static boolean isLineNo(String lineNo) {
        return lineNoList.contains(lineNo);
    }

    /**
     * 判断并保存线路的主线编号
     *
     * @param lineName 线路名称
     * @return 已经存在对应的线路主线编号
     */
    public static String getLineParentNo(String lineName) {
        if (lineParentNoList.containsKey(lineName)) {
            return lineParentNoList.get(lineName);
        }
        return genLineParentNo(lineName);
    }

    /**
     * 生成主线路编号
     *
     * @param lineName
     * @return
     */
    public synchronized static String genLineParentNo(String lineName) {
        String lineNo = isLineParentNo(new Date().getTime() + "P");
        lineParentNoList.put(lineName, lineNo);
        return lineNo;
    }

    /**
     * 判断线路的主线编号是否存在，若存在则重新生成一个编号
     *
     * @param lineNo 线路编号
     * @return 已经存在对应的线路主线编号
     */
    private static String isLineParentNo(String lineNo) {
        if (lineParentNoList.containsValue(lineNo)) {
            try {
                //线程休眠1MS，避免因在同一时间内执行过多次导致StackOverflowError异常
                Thread.sleep(1);
                return isLineParentNo(System.currentTimeMillis() + "P");
            } catch (InterruptedException e) {
                e.printStackTrace();
                logError(e);
            }
        }
        return lineNo;
    }

    /**
     * 获取变电站标识器编号并判断是否已经存在，若不存在则重新生成
     *
     * @param subNo 变电站编号
     * @return 已经存在对应的变电站编号返回true
     */
    public static String getSubMarkerNo(String subNo) {
        if (subMarkerNoMap.containsKey(subNo)) {
            return subMarkerNoMap.get(subNo);
        }
        return genSubMarkerNo(subNo);
    }

    /**
     * 生成变电站标识器编号
     *
     * @param subNo
     * @return
     */
    public static synchronized String genSubMarkerNo(String subNo) {
        String subMarkerNo = isSubMarkerNo("000-001-" + (int) ((Math.random() * 9000) + 1000));
        subMarkerNoMap.put(subNo, subMarkerNo);
        return subMarkerNo;
    }

    /**
     * 判断变电站标识器编号是否存在，若存在则重新生成一个编号
     *
     * @param subMarkerNo 线路编号
     * @return 已经存在对应的线路主线编号
     */
    private static String isSubMarkerNo(String subMarkerNo) {
        if (subMarkerNoMap.containsValue(subMarkerNo)) {
            //线程休眠1MS，避免因在同一时间内执行过多次导致StackOverflowError异常
            return isSubMarkerNo("000-001-" + (int) ((Math.random() * 9000) + 1000));
        }
        return subMarkerNo;
    }

    public static final String groupName_1 = "主网";
    public static final String groupName_2 = "配网";

    /**
     * 根据参数电压等级判断是否属于主配网
     *
     * @param voltage 电压等级
     * @return 返回主网或者配网
     */
    public static String getGroupName(String voltage) {
        switch (voltage) {
            case "10":
                return groupName_2;
            case "110":
                return groupName_1;
            default:
        }
        return null;
    }
}

