package com.demo.springboot.service;

import com.demo.springboot.controller.PSFS;
import com.demo.springboot.message.MessageException;

import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public abstract class FileManageHandle implements FileManage {

    @Override
    public void clearAll(String path) {
        File file = new File(path);
        if(file.exists()) {
            if(file.isFile()) {
                file.delete();
            } else if (file.isDirectory()) {
                File[] files = file.listFiles();
                for(File f : files) {
                    clearAll(f.getAbsolutePath());
                }
                file.delete();
            }
        }
    }

    public static String exists(String path) {
        File tempFile = new File(path);
        if(!tempFile.exists()) {
            tempFile.mkdirs();
        }
        return path;
    }

    @Override
    public void compress (String srcPath, String destPath) {
//    public void compress () {
        File srcFile = new File(srcPath);
        if(!srcFile.exists()) {
            throw new MessageException("文件路径不存在");
        }
        File zipFile = new File(destPath);
        String baseDir = "";
        try {
            ZipOutputStream zouts = new ZipOutputStream(new FileOutputStream(zipFile));
            compressByType(srcFile, zouts, baseDir);
            zouts.finish();
            zouts.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public void compressByType(File srcFile, ZipOutputStream zouts, String baseDir) {
        if(!srcFile.exists()) {
            throw new RuntimeException("文件路径不存在");
        }
        if(srcFile.isFile()) {
            compressFile(srcFile, zouts, baseDir);
        } else if(srcFile.isDirectory()) {
            compressDirectory(srcFile, zouts, baseDir);
        }
    }

    public void compressFile(File srcFile, ZipOutputStream zouts, String baseDir) {
        if(!srcFile.exists()) {
            throw new RuntimeException("文件路径不存在");
        }
        try {
            BufferedInputStream bins = new BufferedInputStream(new FileInputStream(srcFile));
            ZipEntry zipEntry = new ZipEntry(srcFile.getName());
            zouts.putNextEntry(zipEntry);
            int count;
            byte[] buf = new byte[1024];
            while((count = bins.read(buf)) != -1) {
                zouts.write(buf, 0, count);
            }
            bins.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void compressDirectory(File srcFile, ZipOutputStream zouts, String baseDir) {
        if (!srcFile.exists())
            return;
        File[] files = srcFile.listFiles();
        if(files.length == 0){
            try {
                zouts.putNextEntry(new ZipEntry(baseDir + srcFile.getName() + File.separator));
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        for (File file : files) {
            compressByType(file, zouts, baseDir + srcFile.getName() + File.separator);
        }
    }

}
