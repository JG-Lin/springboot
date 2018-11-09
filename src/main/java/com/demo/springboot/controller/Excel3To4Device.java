package com.demo.springboot.controller;

import com.demo.springboot.service.Excel3To4Util;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by csh on 2018/8/2.
 */
public class Excel3To4Device extends Excel3To4Util {
    public static final Map<String, String[]> titleMap = new HashMap<String, String[]>() {
        {
            put("杆塔", new String[]{"系统编号", "标识器编号","经度","纬度","创建时间", "更新时间", "杆塔编号", "维护部门", "运行状态", "产权性质", "投运日期", "生产厂家", "设备型号", "出厂编号", "出厂日期", "杆塔材质", "转角度数", "杆塔全高（m）", "备注"});
            put("标签", new String[]{"系统编号", "线路名", "线路编号", "标识器编号","经度","纬度", "标签编号", "绑扎对象[电缆本体、中间接头、终端设备]", "安装序号", "与上一标签距离", "设备名称", "设备类型", "设备系统编号", "最近工井方向", "距最近工井距离", "附属设施情况", "位置地图", "周围通道分布情况", "通道内电缆情况", "地标路口信息", "对象型号", "对象长度"});
            put("中间头", new String[]{"系统编号", "线路编号", "标识器编号","经度","纬度", "设备名称", "设备型号", "工艺特征[冷缩式、热缩式、预制式、充油式]", "生产厂家", "生产日期", "投运日期", "安装单位", "施工员", "类型[绝缘接头、塞止接头、直通接头、过渡接头、其他]", "厂家联系方式", "中间接头故障情况", "上次巡检日期", "电缆载流量", "中间接头温度", "接地电流情况"});
            put("电房", new String[]{"系统编号", "标识器编号","经度","纬度", "运行编号", "铭牌ID", "电压等级", "所属地市", "运维单位", "维护班组", "设备状态", "投运日期", "是否农网", "是否代维", "配变台数", "配变总容量", "无功补偿容量", "防误方式|机械闭锁,电气闭锁,微机五防", "是否独立建筑物", "是否地下站", "接地电阻(Ω)", "站址", "地区特征", "重要程度", "工程编号", "工程名称", "资产性质", "资产单位", "采集时间", "备注"});
            put("环网柜", new String[]{"系统编号", "线路编号", "标识器编号","经度","纬度", "设备名称", "运行编号", "电系铭牌ID", "电压等级", "运维单位", "维护班组", "设备状态[库存备用、现场留用、未投运、在运、退役、待报废、报废]", "是否代维[是、否]", "是否农网[是、否]", "型号", "生产厂家", "出厂编号", "出厂日期", "投运日期", "接地电阻(Ω)", "备用进出线间隔数", "站址", "地区特征[林区、牧区、郊区、县城、市区、农村、市中心区]", "重要程度[一般、特别重要、重要]", "资产性质[国家电网公司、南方电网公司]", "资产单位", "工程编号", "工程名称", "设备增加方式[融资租入、盘盈、零星购置、其他途径、基本建设、投资者投入、无偿调入、接受捐赠、债务重组取得、技术改造]", "备注"});
            put("接线箱", new String[]{"系统编号", "线路编号", "标识器编号","经度","纬度", "电压等级", "设备名称", "运行编号", "电系铭牌ID", "运维单位", "维护班组", "设备状态[库存备用、现场留用、未投运、在运、退役、待报废、报废]", "是否代维[是、否]", "是否农网[是、否]", "类型[电缆分支箱、电缆分界室]", "型号", "生产厂家", "出厂编号", "出厂日期", "投运日期", "接地电阻(Ω)", "站址", "地区特征[林区、牧区、郊区、县城、市区、农村、市中心区]", "重要程度[一般、特别重要、重要]", "资产性质[国家电网公司、南方电网公司]", "资产单位", "工程编号", "工程名称", "设备增加方式[融资租入、盘盈、零星购置、其他途径、基本建设、投资者投入、无偿调入、接受捐赠、债务重组取得、技术改造]", "备注"});
            put("工井", new String[]{"系统编号", "标识器编号","经度","纬度", "断面信息", "工程名称", "盖板厚度(cm)", "盖板块数", "所属项目", "所属机构", "运维单位", "维护班组", "设备名称", "所属责任区", "工井形状", "所属地市", "地区特征", "地理位置", "井位置", "井类型|直线井,接头井,转角井,三通井,四通井,顶管工井", "结构|带室,不带室", "井面高程(m)", "内底高程(m)", "井盖形状|方形,圆形", "井盖尺寸(m)", "井盖材质|球墨,铁,水泥,复合材料,聚酯,其他", "井盖生产厂家", "井盖出厂日期", "平台层数", "施工单位", "施工日期", "峻工日期", "图纸编号", "资产性质|国家电网公司,分部,省（直辖市、自治区）公司,子公司,用户", "资产单位", "资产编号", "专业分类", "备注", "采集时间"});
        }
    };

    private File book3File;

    private Map<String, Device> deviceMap = new HashMap();

    public void execute(String path3, String path4) {
        init(path3, path4);
        transform();
    }

    /**
     * 初始化所有设备对象，并且初始化excel表头
     *
     * @param path3
     * @param path4
     */
    private void init(String path3, String path4) {

        try {
            String[] devices = new String[titleMap.keySet().size()];
            titleMap.keySet().toArray(devices);
            book3File = new File(path3);
            //初始化所有设备对象，excel表头等
            for (int i = 0; i < devices.length; i++) {
                Device device = new Device(path4 + devices[i] + xlsx);
                Workbook book_4 = device.getBook_4();
                Sheet sheet_4 = book_4.createSheet(devices[i]);
                Row row_4 = sheet_4.createRow(0);
                // 初始化标题
                String[] title = titleMap.get(devices[i]);
                for (int j = 0; j < title.length; j++) {
                    Cell cell_4 = row_4.createCell(j);
                    cell_4.setCellValue(title[j]);
                }
                deviceMap.put(devices[i], device);
            }
        } catch (Exception e) {
            Excel3To4DataCache.logError(e);
            e.printStackTrace();
        }
    }

    /**
     * 数据转化
     */
    private void transform() {
        try {
            File[] File3 = book3File.listFiles();
            if (File3 != null) {
                for (File file : File3) {
                    Workbook book_3 = getExcelWordbook(file);
                    Sheet sheet = book_3.getSheetAt(0);
                    int lastRowNum = sheet.getLastRowNum();
                    for (int i = 1; i <= lastRowNum; i++) {
                        Row row_3 = sheet.getRow(i);
                        String type = row_3 == null ? "" : getCellValue(row_3.getCell(D));
                        switch (type) {
                            case "杆塔":
                                executeGT(row_3);
                                break;
                            case "标签":
                                executeBQ(row_3);
                                break;
                            case "中间头":
                                executeZJT(row_3);
                                break;
                            case "电房":
                                executeDF(row_3);
                                break;
                            case "环网柜":
                                executeHWG(row_3);
                                break;
                            case "接线箱":
                                executeJXX(row_3);
                                break;
                            case "工井":
                                executeGJ(row_3);
                                break;
                        }
                    }
                }
            }
            String[] devices = new String[titleMap.keySet().size()];
            titleMap.keySet().toArray(devices);
            for (int i = 0; i < devices.length; i++) {

                outFile(devices[i]);
            }
        } catch (Exception e) {
            System.out.println("3.0设备转化4.0设备失败");
            e.printStackTrace();
            Excel3To4DataCache.logError(e);
        }
    }

    /**
     * 输出excel
     *
     * @param type
     */
    private void outFile(String type) {
        Device device = deviceMap.get(type);
        try {
            device.getBook_4().write(device.getBook4_FileOut());

            int num = device.getBook_4().getSheetAt(0).getLastRowNum();

            if(num <= 1) {
                File file = new File(device.getFilePath());

                file.delete();
            }
        } catch (IOException e) {
            Excel3To4DataCache.logError(e);
            e.printStackTrace();
        } finally {
            try {
                device.getBook4_FileOut().flush();
                device.getBook4_FileOut().close();
            } catch (IOException e) {
                Excel3To4DataCache.logError(e);
                e.printStackTrace();
            }
        }
    }

    /**
     * 处理杆塔
     *
     * @param row
     */
    private void executeGT(Row row) {
        Device device = deviceMap.get("杆塔");
        Workbook book4 = device.getBook_4();
        Sheet sheet = book4.getSheetAt(0);
        Row row4 = sheet.createRow(sheet.getLastRowNum() + 1);
        //标识器编号
        Cell cell4 = row4.createCell(B);
        cell4.setCellValue(getCellValue(row.getCell(A)));
        //经度
        cell4 = row4.createCell(C);
        cell4.setCellValue(getCellValue(row.getCell(E)));
        //纬度
        cell4 = row4.createCell(D);
        cell4.setCellValue(getCellValue(row.getCell(F)));
        //杆塔编号
        cell4 = row4.createCell(G);
        cell4.setCellValue(getCellValue(row.getCell(B)));
        //生产厂家
        cell4 = row4.createCell(L);
        cell4.setCellValue(getCellValue(row.getCell(P)));
        //设备型号
        cell4 = row4.createCell(M);
        cell4.setCellValue(getCellValue(row.getCell(O)));
        //出厂日期
        cell4 = row4.createCell(O);
        cell4.setCellValue(getDateCell(row.getCell(Q)));
    }

    /**
     * 处理标签
     *
     * @param row
     */
    private void executeBQ(Row row) {
        Device device = deviceMap.get("标签");
        Workbook book4 = device.getBook_4();
        Sheet sheet = book4.getSheetAt(0);
        Row row4 = sheet.createRow(sheet.getLastRowNum() + 1);
        //线路名
        Cell cell4 = row4.createCell(B);
        cell4.setCellValue(getCellValue(row.getCell(K)));
        //线路编号
        cell4 = row4.createCell(C);
        cell4.setCellValue(getCellValue(row.getCell(L)));
        //标识器编号
        cell4 = row4.createCell(D);
        cell4.setCellValue(getCellValue(row.getCell(A)));
        //经度
        cell4 = row4.createCell(E);
        cell4.setCellValue(getCellValue(row.getCell(E)));
        //纬度
        cell4 = row4.createCell(F);
        cell4.setCellValue(getCellValue(row.getCell(F)));


        //标签编号
        cell4 = row4.createCell(G);
        cell4.setCellValue(getCellValue(row.getCell(B)));
        //绑扎对象[电缆本体、中间接头、终端设备]
        cell4 = row4.createCell(H);
        cell4.setCellValue(getCellValue(row.getCell(C)));
        //设备名称
        cell4 = row4.createCell(K);
        cell4.setCellValue(getCellValue(row.getCell(B)));
    }

    /**
     * 处理中间头
     *
     * @param row
     */
    private void executeZJT(Row row) {
        Device device = deviceMap.get("中间头");
        Workbook book4 = device.getBook_4();
        Sheet sheet = book4.getSheetAt(0);
        Row row4 = sheet.createRow(sheet.getLastRowNum() + 1);
        //线路编号
        Cell cell4 = row4.createCell(B);
        cell4.setCellValue(getCellValue(row.getCell(L)));
        //标识器编号
        cell4 = row4.createCell(C);
        cell4.setCellValue(getCellValue(row.getCell(A)));
        //经度
        cell4 = row4.createCell(D);
        cell4.setCellValue(getCellValue(row.getCell(E)));
        //纬度
        cell4 = row4.createCell(E);
        cell4.setCellValue(getCellValue(row.getCell(F)));
        //设备型号
        cell4 = row4.createCell(G);
        cell4.setCellValue(getCellValue(row.getCell(O)));
        //生产厂家
        cell4 = row4.createCell(I);
        cell4.setCellValue(getCellValue(row.getCell(P)));
        //生产日期
        cell4 = row4.createCell(J);
        cell4.setCellValue(getDateCell(row.getCell(Q)));
        //安装单位
        cell4 = row4.createCell(L);
        cell4.setCellValue(getCellValue(row.getCell(R)));
        //施工员
        cell4 = row4.createCell(M);
        cell4.setCellValue(getCellValue(row.getCell(S)));
        //设备名称
        cell4 = row4.createCell(F);
        cell4.setCellValue(getCellValue(row.getCell(B)));
    }

    /**
     * 处理电房
     *
     * @param row
     */
    private void executeDF(Row row) {
        Device device = deviceMap.get("电房");
        Workbook book4 = device.getBook_4();
        Sheet sheet = book4.getSheetAt(0);
        Row row4 = sheet.createRow(sheet.getLastRowNum() + 1);
        //标识器编号
        Cell cell4 = row4.createCell(B);
        cell4.setCellValue(getCellValue(row.getCell(A)));
        //经度
        cell4 = row4.createCell(C);
        cell4.setCellValue(getCellValue(row.getCell(E)));
        //纬度
        cell4 = row4.createCell(D);
        cell4.setCellValue(getCellValue(row.getCell(F)));

    }

    /**
     * 处理环网柜
     *
     * @param row
     */
    private void executeHWG(Row row) {
        Device device = deviceMap.get("环网柜");
        Workbook book4 = device.getBook_4();
        Sheet sheet = book4.getSheetAt(0);
        Row row4 = sheet.createRow(sheet.getLastRowNum() + 1);
        //线路编号
        Cell cell4 = row4.createCell(B);
        cell4.setCellValue(getCellValue(row.getCell(L)));
        //标识器编号
        cell4 = row4.createCell(C);
        cell4.setCellValue(getCellValue(row.getCell(A)));
        //经度
        cell4 = row4.createCell(D);
        cell4.setCellValue(getCellValue(row.getCell(E)));
        //纬度
        cell4 = row4.createCell(E);
        cell4.setCellValue(getCellValue(row.getCell(F)));
        //型号
        cell4 = row4.createCell(O);
        cell4.setCellValue(getCellValue(row.getCell(O)));
        //生产厂家
        cell4 = row4.createCell(P);
        cell4.setCellValue(getCellValue(row.getCell(P)));
        //出厂日期
        cell4 = row4.createCell(R);
        cell4.setCellValue(getDateCell(row.getCell(Q)));
        //工程名称
        cell4 = row4.createCell(AB);
        cell4.setCellValue(getCellValue(row.getCell(M)));
    }

    /**
     * 处理接线箱
     *
     * @param row
     */
    private void executeJXX(Row row) {
        Device device = deviceMap.get("接线箱");
        Workbook book4 = device.getBook_4();
        Sheet sheet = book4.getSheetAt(0);
        Row row4 = sheet.createRow(sheet.getLastRowNum() + 1);
        //线路编号
        Cell cell4 = row4.createCell(B);
        cell4.setCellValue(getCellValue(row.getCell(L)));
        //标识器编号
        cell4 = row4.createCell(C);
        cell4.setCellValue(getCellValue(row.getCell(A)));
        //经度
        cell4 = row4.createCell(D);
        cell4.setCellValue(getCellValue(row.getCell(E)));
        //纬度
        cell4 = row4.createCell(E);
        cell4.setCellValue(getCellValue(row.getCell(F)));
        //型号
        cell4 = row4.createCell(P);
        cell4.setCellValue(getCellValue(row.getCell(O)));
        //生产厂家
        cell4 = row4.createCell(Q);
        cell4.setCellValue(getCellValue(row.getCell(P)));
        //出厂日期
        cell4 = row4.createCell(S);
        cell4.setCellValue(getDateCell(row.getCell(Q)));
        //工程名称
        cell4 = row4.createCell(AB);
        cell4.setCellValue(getCellValue(row.getCell(M)));
    }

    /**
     * 处理工井
     *
     * @param row
     */
    private void executeGJ(Row row) {
        Device device = deviceMap.get("工井");
        Workbook book4 = device.getBook_4();
        Sheet sheet = book4.getSheetAt(0);
        Row row4 = sheet.createRow(sheet.getLastRowNum() + 1);
        //标识器编号
        Cell cell4 = row4.createCell(B);
        cell4.setCellValue(getCellValue(row.getCell(A)));
        //经度
        cell4 = row4.createCell(C);
        cell4.setCellValue(getCellValue(row.getCell(E)));
        //纬度
        cell4 = row4.createCell(D);
        cell4.setCellValue(getCellValue(row.getCell(F)));
        //工程名称
        cell4 = row4.createCell(F);
        cell4.setCellValue(getCellValue(row.getCell(M)));
        //井盖生产厂家
        cell4 = row4.createCell(AA);
        cell4.setCellValue(getCellValue(row.getCell(P)));
        //施工单位
        cell4 = row4.createCell(AD);
        cell4.setCellValue(getCellValue(row.getCell(R)));
        //施工日期
        cell4 = row4.createCell(AE);
        cell4.setCellValue(getDateCell(row.getCell(T)));
        //设备名称
        cell4 = row4.createCell(M);
        cell4.setCellValue(getCellValue(row.getCell(B)));
    }
}
