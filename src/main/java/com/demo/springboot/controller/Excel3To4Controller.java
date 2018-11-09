package com.demo.springboot.controller;

import com.demo.springboot.message.HandleResponse;
import com.demo.springboot.message.MessageException;
import com.demo.springboot.service.Excel3To4;
import com.demo.springboot.service.FileManage;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;

@RestController
@RequestMapping("/upload")
public class Excel3To4Controller {

    @Autowired
    private Excel3To4 excel3To4;
    @Autowired
    private FileManage fileManage;

    @PostMapping(value = "/e3to4")
    @ResponseBody
    public Object fromExcel3(HttpServletRequest request, HttpServletResponse response) {
        MultipartHttpServletRequest mRequest = (MultipartHttpServletRequest)request;
        try {
            List<MultipartFile> subFiles = mRequest.getFiles("sub");
            ifFilePostfix(subFiles);
            List<MultipartFile> lineFiles = mRequest.getFiles("line");
            ifFilePostfix(lineFiles);
            List<MultipartFile> nodeFiles = mRequest.getFiles("node");
            ifFilePostfix(nodeFiles);
            List<MultipartFile> devFiles = mRequest.getFiles("dev");
            fileManage.writer(subFiles, "变电站");
            fileManage.writer(lineFiles, "线路信息");
            fileManage.writer(nodeFiles, "标识器");
            fileManage.writer(devFiles, "设备");
            excel3To4.execute();
            fileManage.compress(PSFS.ABSOLUTE_PATH + "\\4.0", PSFS.ABSOLUTE_PATH + "\\4.0.zip");


            File file = new File(PSFS.ABSOLUTE_PATH + "\\4.0.zip");
            response.setHeader("content-type", "application/octet-stream");
            response.setContentType("application/octet-stream");
            response.setHeader("Content-Disposition", "attachment;filename=4.0.zip");
            try {
                OutputStream outputStream = response.getOutputStream();
                BufferedInputStream bins = new BufferedInputStream(new FileInputStream(file));
                byte[] buf = new byte[1024];
                while((bins.read(buf)) != -1) {
                    outputStream.write(buf);
                }
                outputStream.flush();
                bins.close();
                outputStream.close();
                fileManage.clearAll(PSFS.ABSOLUTE_PATH);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (MessageException e) {
            e.printStackTrace();
            return HandleResponse.returnMessage(-1,e.getInfo());
        }
        return HandleResponse.returnMessage(0);
    }

    private void ifFilePostfix(List<MultipartFile> files) throws MessageException {
        if(null == files || files.size() == 0) {
            throw new MessageException("没有上传文件,后者漏上传");
        }
        for (MultipartFile f : files) {
            if(!f.getOriginalFilename().endsWith(".xlsx") && !f.getOriginalFilename().endsWith(".xls")) {
                throw new MessageException(f.getOriginalFilename() + "类型不对");
            }
        }
    }

}
