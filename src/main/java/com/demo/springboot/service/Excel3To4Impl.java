package com.demo.springboot.service;

import com.demo.springboot.controller.Excel3To4DataCache;
import com.demo.springboot.controller.Excel3To4Device;
import com.demo.springboot.controller.PSFS;
import com.demo.springboot.controller.Substation;
import com.demo.springboot.message.MessageException;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * FileName: Excel3To4Impl
 * Author:   LK
 * Date:     2018/7/28 9:55
 * Description: 1.0.0
 */
@Service
public class Excel3To4Impl extends Excel3To4Util implements Excel3To4 {
    public File sub_File3_0;
    public File sub_File4_0;
    public File line_File3_0;
    public File line_File4_0;
    public File node_File3_0;
    public File node_File4_0;
    public File lineNode_File4_0;
    private String absolutePath = PSFS.ABSOLUTE_PATH;

    /**
     * 保存变电站标识器编号
     */
    public Map<String, Substation> subMakerNoMap = new ConcurrentHashMap<>();

    @Override
    public void execute() {
        sub_File3_0 = new File(absolutePath + "\\3.0\\变电站\\");
        sub_File4_0 = new File(FileManageHandle.exists(absolutePath + "\\4.0") + "\\变电站.xlsx");
        line_File3_0 = new File(absolutePath + "\\3.0\\线路信息\\");
        line_File4_0 = new File(FileManageHandle.exists(absolutePath + "\\4.0") + "\\线路.xlsx");
        node_File3_0 = new File(absolutePath + "\\3.0\\标识器\\");
        node_File4_0 = new File(FileManageHandle.exists(absolutePath + "\\4.0") + "\\标识器.xlsx");
        lineNode_File4_0 = new File(absolutePath + "\\4.0\\线路节点.xlsx");
//        System.out.println("开始转" + this.projectName + "变电站。。。");
        fromSubExcel3();
//        System.out.println("完成"+ this.projectName +"变电站");
//        System.out.println("开始转"+ this.projectName +"线路。。。");
        fromLineExcel3();
//        System.out.println("完成"+ this.projectName +"线路");
//        System.out.println("开始转"+ this.projectName +"标识器，线路节点。。。");
        fromNodeExcel3();
//        System.out.println("完成项目：" + this.projectName);
        Excel3To4Device excel3To4Device = new Excel3To4Device();
        excel3To4Device.execute(FileManageHandle.exists(absolutePath + "\\3.0\\设备"), FileManageHandle.exists(absolutePath + "\\4.0"));
    }

    /**
     * 读取3.0的变电站excel表，转成4.0变电站表
     */
    public void fromSubExcel3() {
        try {
            // 新建4.0变电站excel表
            Workbook sub_book_4 = new XSSFWorkbook();
            Sheet sub_sheet_4 = sub_book_4.createSheet("变电站");
            Row sub_row_4 = sub_sheet_4.createRow(0);
            // 初始化变电站标题
            for (int i = 0; i < subTitle.length; i++) {
                Cell sub_cell_4 = sub_row_4.createCell(i);
                sub_cell_4.setCellValue(subTitle[i]);
            }
            File[] subs_File3_0 = sub_File3_0.listFiles();
            int row4 = sub_sheet_4.getLastRowNum();
            for (File sub_File : subs_File3_0) {
                Workbook sub_book_3 = getExcelWordbook(sub_File);
                Sheet sheet = sub_book_3.getSheetAt(0);
                int lastRowNum = sheet.getLastRowNum();
                for (int i = 1; i <= lastRowNum; i++) {
                    Row sub_row_3 = sheet.getRow(i);
                    //判断当前行是否为空，并且判断变电站编号是否为空
                    if (sub_row_3 == null || StringUtils.isEmpty(getCellValue(sub_row_3.getCell(B)))) {
                        continue;
                    }
                    Substation substation = new Substation();
                    row4++;
                    sub_row_4 = sub_sheet_4.createRow(row4);
                    // 变电站名称
                    Cell sub_cell_4 = sub_row_4.createCell(A);
                    sub_cell_4.setCellValue(getCellValue(sub_row_3.getCell(C)));
                    // 电压等级
                    sub_cell_4 = sub_row_4.createCell(B);
                    sub_cell_4.setCellValue(getCellValue(sub_row_3.getCell(A)));
                    // 变电站编号
                    sub_cell_4 = sub_row_4.createCell(C);
                    String subNo = getCellValue(sub_row_3.getCell(B));
                    sub_cell_4.setCellValue(subNo);
                    // 标识器编号
                    sub_cell_4 = sub_row_4.createCell(D);
                    String subMarkerNo = Excel3To4DataCache.getSubMarkerNo(subNo);
                    sub_cell_4.setCellValue(subMarkerNo);

                    //经度
                    sub_cell_4 = sub_row_4.createCell(E);
                    sub_cell_4.setCellValue(getCellValue(sub_row_3.getCell(D)));
                    //纬度
                    sub_cell_4 = sub_row_4.createCell(F);
                    sub_cell_4.setCellValue(getCellValue(sub_row_3.getCell(E)));
//                sub_cell_4.setCellValue(getCellValue(sub_row_3.getCell(A)));
                    // 产权性质(0公用、1 专用)
                    sub_cell_4 = sub_row_4.createCell(G);
                    sub_cell_4.setCellValue("公用");
                    // 采集时间
//                sub_cell_4 = sub_row_4.createCell(F);
                    // 记录变电站编号，用来验证转线路中变电站编号
                    substation.setSubNo(subNo);
                    substation.setMarkerNo(subMarkerNo);
                    substation.setLongitude(getCellValue(sub_row_3.getCell(D)));
                    substation.setLatitude(getCellValue(sub_row_3.getCell(E)));
                    subMakerNoMap.put(subNo, substation);
                    Excel3To4DataCache.subNoList.add(subNo);
                }
            }
            FileOutputStream sub4_FileOut = new FileOutputStream(sub_File4_0);
            sub_book_4.write(sub4_FileOut);
            sub4_FileOut.flush();
            sub4_FileOut.close();
        } catch (Exception e) {
//            Excel3To4DataCache.logError(e);
            e.printStackTrace();
            throw new MessageException("变电站3.0转4.0失败");
        }
    }

    /**
     * 读取3.0的线路excel表，转成4.0线路表
     */
    public void fromLineExcel3() {
        try {
            // 创建线路excel表
            Workbook line_book_4 = new XSSFWorkbook();
            Sheet line_sheet_4 = line_book_4.createSheet("线路");
            Row line_row_4 = line_sheet_4.createRow(0);
            // 初始化线路标题
            for (int i = 0; i < lineTitle.length; i++) {
                Cell line_cell_4 = line_row_4.createCell(i);
                line_cell_4.setCellValue(lineTitle[i]);
            }
            File[] files = line_File3_0.listFiles();
            int row4 = line_sheet_4.getLastRowNum();
            for (File line_File : files) {
                Workbook line_book_3 = getExcelWordbook(line_File);
                Sheet line_sheet_3 = line_book_3.getSheetAt(0);
                int lastRowNum = line_sheet_3.getLastRowNum();
                for (int i = 1; i <= lastRowNum; i++) {
                    Row line_row_3 = line_sheet_3.getRow(i);
                    //判断当前行是否为空，并且判断线路编号是否为空
                    if (line_row_3 == null || StringUtils.isEmpty(getCellValue(line_row_3.getCell(E)))) {
                        continue;
                    }
                    row4++;
                    line_row_4 = line_sheet_4.createRow(row4);
                    // 线路名称
                    Cell line_cell_4 = line_row_4.createCell(A);
                    line_cell_4.setCellValue(getCellValue(line_row_3.getCell(C)));
                    // 线路编号
                    line_cell_4 = line_row_4.createCell(B);
                    line_cell_4.setCellValue(getCellValue(line_row_3.getCell(B)));
                    // 主线名称
                    line_cell_4 = line_row_4.createCell(C);
                    String lineParentName = getCellValue(line_row_3.getCell(I));
                    line_cell_4.setCellValue(lineParentName);
                    // 主线路编号
                    line_cell_4 = line_row_4.createCell(D);
                    line_cell_4.setCellValue(Excel3To4DataCache.getLineParentNo(lineParentName));
                    // 电压等级
                    line_cell_4 = line_row_4.createCell(E);
                    String voltage = getCellValue(line_row_3.getCell(A));
                    line_cell_4.setCellValue(voltage);
                    // 所属变电站
                    line_cell_4 = line_row_4.createCell(F);
                    String subNo = getCellValue(line_row_3.getCell(D));
                    if (!Excel3To4DataCache.isSubNo(subNo)) {
                        throw new MessageException(line_File.getName() + "第" + (i + 1) + "行中变电站编号不存在：" + subNo);
                    }
                    line_cell_4.setCellValue(subNo);
                    // 结束变电站
                    line_cell_4 = line_row_4.createCell(G);
                    line_cell_4.setCellValue(getCellValue(line_row_3.getCell(E)));
                    // 线路偏移
                    line_cell_4 = line_row_4.createCell(H);
                    line_cell_4.setCellValue(getCellValue(line_row_3.getCell(H)));
                    // 线路分组
                    line_cell_4 = line_row_4.createCell(I);
                    line_cell_4.setCellValue(Excel3To4DataCache.getGroupName(voltage));
                    // 维护部门
                    // 设备型号
                    // 产权性质
                    // 生产厂家
                    // 出厂编号
                    // 投运日期
                    // 出厂日期
                    // 图上长度(m)
                    // 实测长度(m)
                    // 敷设方式
                    // 备注
                    // 序号
                    // 线路类型(0:电缆工程线路、1:工程施工的线路段)
                    // 导线类型
                    Excel3To4DataCache.lineNoList.add(getCellValue(line_row_3.getCell(B)));
                }
            }
            FileOutputStream line4_FileOut = new FileOutputStream(line_File4_0);
            line_book_4.write(line4_FileOut);
            line4_FileOut.flush();
            line4_FileOut.close();
        } catch (Exception e) {
//            Excel3To4DataCache.logError(e);
            e.printStackTrace();
            throw new MessageException("线路站3.0转4.0失败");
        }
    }

    /**
     * 读取3.0的标识器excel表，转成4.0的标识器跟线路节点表
     */
    public void fromNodeExcel3() {
        try {
            // 创建标识器表
            Workbook node_book_4 = new XSSFWorkbook();
            Sheet node_sheet_4 = node_book_4.createSheet("标识器");
            Row node_row_4 = node_sheet_4.createRow(0);
            // 初始化标题
            for (int i = 0; i < nodeTitle.length; i++) {
                Cell cell = node_row_4.createCell(i);
                cell.setCellValue(nodeTitle[i]);
            }
            int node_row4 = node_sheet_4.getLastRowNum();
            // 加入变电站路径点
            for (Map.Entry<String, Substation> entry : subMakerNoMap.entrySet()) {
                node_row4++;
                node_row_4 = node_sheet_4.createRow(node_row4);
                // 标识器编号
                Cell node_cell_4 = node_row_4.createCell(A);
                node_cell_4.setCellValue(entry.getValue().getMarkerNo());
                // 标识器类型(标识球、 标识钉、杆塔、 路径点)
                node_cell_4 = node_row_4.createCell(B);
                node_cell_4.setCellValue("虚拟点");
                // 经度
                node_cell_4 = node_row_4.createCell(D);
                node_cell_4.setCellValue(entry.getValue().getLongitude());
                // 纬度
                node_cell_4 = node_row_4.createCell(E);
                node_cell_4.setCellValue(entry.getValue().getLatitude());
                // 备注
                node_cell_4 = node_row_4.createCell(Q);
                node_cell_4.setCellValue("此路径点为变电站，编号为："+entry.getValue().getSubNo());
            }
            // 创建线路节点表
            Workbook lineNode_book_4 = new XSSFWorkbook();
            Sheet lineNode_sheet_4 = lineNode_book_4.createSheet("线路节点");
            Row lineNode_row_4 = lineNode_sheet_4.createRow(0);
            // 初始化标题
            for (int i = 0; i < lineNodeTitle.length; i++) {
                Cell cell = lineNode_row_4.createCell(i);
                cell.setCellValue(lineNodeTitle[i]);
            }
            File[] files = node_File3_0.listFiles();
            int line_row4 = lineNode_sheet_4.getLastRowNum();
            node_row4 = node_sheet_4.getLastRowNum();
            for (File file : files) {
                Workbook node_book_3 = getExcelWordbook(file);
                Sheet node_sheet_3 = node_book_3.getSheetAt(0);
                int lastRowNum = node_sheet_3.getLastRowNum();
                for (int i = 1; i <= lastRowNum; i++) {
                    Row node_row_3 = node_sheet_3.getRow(i);
                    //判断当前行是否为空，并且判断线路编号是否为空
                    if (node_row_3 == null || StringUtils.isEmpty(getCellValue(node_row_3.getCell(C)))) {
                        continue;
                    }
                    line_row4++;
                    lineNode_row_4 = lineNode_sheet_4.createRow(line_row4);
                    // 线路名称
                    Cell lineNode_cell_4 = lineNode_row_4.createCell(A);
                    lineNode_cell_4.setCellValue(getCellValue(node_row_3.getCell(D)));
                    // 线路编号
                    lineNode_cell_4 = lineNode_row_4.createCell(B);
                    String lineNo = getCellValue(node_row_3.getCell(C));
                    if (!Excel3To4DataCache.isLineNo(lineNo)) {
                        throw new MessageException(file.getName() + "第" + (i + 1) + "行中线路编号不存在：" + lineNo);
                    }
                    lineNode_cell_4.setCellValue(lineNo);
                    //标识器编号
                    lineNode_cell_4 = lineNode_row_4.createCell(C);
                    lineNode_cell_4.setCellValue(getCellValue(node_row_3.getCell(A)));
                    //经度
                    lineNode_cell_4 = lineNode_row_4.createCell(D);
                    lineNode_cell_4.setCellValue(getCellValue(node_row_3.getCell(H)));
                    //纬度
                    lineNode_cell_4 = lineNode_row_4.createCell(E);
                    lineNode_cell_4.setCellValue(getCellValue(node_row_3.getCell(I)));
                    // 连接顺序
                    lineNode_cell_4 = lineNode_row_4.createCell(F);
                    lineNode_cell_4.setCellValue(getCellValue(node_row_3.getCell(F)));
                    // 上一节点距离
                    lineNode_cell_4 = lineNode_row_4.createCell(G);
                    lineNode_cell_4.setCellValue(getCellValue(node_row_3.getCell(Q)));
                    // 上一节点敷设方式
                    lineNode_cell_4 = lineNode_row_4.createCell(H);
                    lineNode_cell_4.setCellValue(getCellValue(node_row_3.getCell(U)));
                    // 深度(距电缆沟、槽、顶管底深度)
                    lineNode_cell_4 = lineNode_row_4.createCell(I);
                    lineNode_cell_4.setCellValue(getCellValue(node_row_3.getCell(N)));
                    // 设备描述
                    lineNode_cell_4 = lineNode_row_4.createCell(J);
                    lineNode_cell_4.setCellValue(getCellValue(node_row_3.getCell(P)));
                    // 备注
//                lineNode_cell_4 = lineNode_row_4.createCell(A);

                    node_row4++;
                    node_row_4 = node_sheet_4.createRow(node_row4);
                    // 标识器编号
                    Cell node_cell_4 = node_row_4.createCell(A);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(A)));
                    // 标识器类型(标识球、 标识钉、杆塔、 路径点)
                    node_cell_4 = node_row_4.createCell(B);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(AA)));
                    // 敷设方式(架空、电缆沟、电缆槽、埋管、顶管)
                    node_cell_4 = node_row_4.createCell(C);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(G)));
                    // 经度
                    node_cell_4 = node_row_4.createCell(D);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(H)));
                    // 纬度
                    node_cell_4 = node_row_4.createCell(E);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(I)));
                    // 地面高程
//                node_cell_4 = node_row_4.createCell(F);

                    // 安装位置
                    node_cell_4 = node_row_4.createCell(G);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(L)));
                    // 地理位置
                    node_cell_4 = node_row_4.createCell(H);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(M)));
                    // 距底深度(电缆沟、槽、顶管)[cm]
                    node_cell_4 = node_row_4.createCell(I);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(N)));
                    // 电缆沟/槽宽度[cm]
                    node_cell_4 = node_row_4.createCell(J);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(O)));
                    // 设备描述
                    node_cell_4 = node_row_4.createCell(K);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(P)));
                    // 同沟电缆回路数[回]
                    node_cell_4 = node_row_4.createCell(L);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(R)));
                    // 同沟电缆名称及相序位置
                    node_cell_4 = node_row_4.createCell(M);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(T)));
                    // 安装图
                    node_cell_4 = node_row_4.createCell(N);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(X)));
                    // 背景图
                    node_cell_4 = node_row_4.createCell(O);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(W)));
                    // 放置日期
                    node_cell_4 = node_row_4.createCell(P);
                    node_cell_4.setCellValue(getDateCell(node_row_3.getCell(Y)));
                    // 备注
                    node_cell_4 = node_row_4.createCell(Q);
                    node_cell_4.setCellValue(getCellValue(node_row_3.getCell(V)));
                    // 坐标编号(用于匹配外业采集的坐标)
//                node_cell_4 = node_row_4.createCell(R);
                }
            }
            FileOutputStream node4_fileOut = new FileOutputStream(node_File4_0);
            node_book_4.write(node4_fileOut);
            node4_fileOut.flush();
            node4_fileOut.close();
            FileOutputStream lineNode4_fileOut = new FileOutputStream(lineNode_File4_0);
            lineNode_book_4.write(lineNode4_fileOut);
            lineNode4_fileOut.flush();
            lineNode4_fileOut.close();
        } catch (Exception e) {
//            Excel3To4DataCache.logError(e);
            e.printStackTrace();
            throw new MessageException("标识器3.0转4.0失败");
        }
    }

}