package com.demo.springboot.controller;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * Created by csh on 2018/8/3.
 */
public class Device {

    private Workbook book_4;

    private String filePath;

    private FileOutputStream book4_FileOut;

    public Device(String book4_FileOut) throws FileNotFoundException {
        this.book4_FileOut = new FileOutputStream(book4_FileOut);
        this.filePath = book4_FileOut;
        this.book_4 = new XSSFWorkbook();
    }

    public Workbook getBook_4() {
        return book_4;
    }

    public void setBook_4(Workbook book_4) {
        this.book_4 = book_4;
    }

    public FileOutputStream getBook4_FileOut() {
        return book4_FileOut;
    }

    public void setBook4_FileOut(FileOutputStream book4_FileOut) {
        this.book4_FileOut = book4_FileOut;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }
}
