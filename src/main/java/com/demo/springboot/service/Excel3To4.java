package com.demo.springboot.service;

import com.demo.springboot.message.MessageException;
import org.springframework.web.multipart.MultipartFile;

import java.io.OutputStream;
import java.util.List;

/**
 * FileName: Excel3To4
 * Author:   LK
 * Date:     2018/7/28 9:55
 * Description: 1.0.0
 */
public interface Excel3To4 {
    /**
     * 执行excel3.0转4.0
     */
    void execute() throws MessageException;

}
