package com.demo.springboot.service;

import com.demo.springboot.controller.PSFS;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.List;

@Service
public class FileManageImpl extends FileManageHandle {

    private String absolutPath = PSFS.ABSOLUTE_PATH;

    @Override
    public void writer(List<MultipartFile> subFiles, String typeName) {
        InputStream ins = null;
        BufferedOutputStream bouts = null;
        byte[] buffer = new byte[1024];
        try {

            for (MultipartFile file : subFiles) {
                String originalFilename = file.getOriginalFilename();
                ins = file.getInputStream();
                bouts = new BufferedOutputStream(new FileOutputStream(new File(exists(absolutPath + "\\3.0\\" + typeName) + "\\" + originalFilename)));
                while ((ins.read(buffer)) != -1) {
                    bouts.write(buffer);
                }
                bouts.flush();
                bouts.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
