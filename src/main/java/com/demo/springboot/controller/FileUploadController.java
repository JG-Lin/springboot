package com.demo.springboot.controller;

import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;

//@RestController
public class FileUploadController {

    @RequestMapping(value = "/upload", method = RequestMethod.POST)
    @ResponseBody
    public Object upload(@RequestParam("file") MultipartFile file) {
        System.out.println(file.getOriginalFilename());
        return "ok";
    }

    @RequestMapping(value = "/uploads", method = RequestMethod.POST)
    @ResponseBody
    public Object uploads(@RequestParam("files") MultipartFile[] files) {
        for ( MultipartFile f : files ) {
            System.out.println(f.getOriginalFilename());
        }
        return "ok";
    }

    @PostMapping(value = "/upload/batch")
    public Object batchUpload(HttpServletRequest request){
        MultipartHttpServletRequest mRequest = (MultipartHttpServletRequest)request;
        List<MultipartFile> files = mRequest.getFiles("files");
        for ( MultipartFile f : files ) {
            System.out.println(f.getOriginalFilename());
        }
        MultipartFile file =  ((MultipartHttpServletRequest)request).getFile("file");
            System.out.println(file.getOriginalFilename());
        return "OKS";
    }

    @GetMapping(value = "/down")
    public Object download(HttpServletResponse response) throws IOException {
//        MultipartHttpServletRequest mRequest = (MultipartHttpServletRequest)request;
//        MultipartFile file = mRequest.getFile("file");
        File file = new File("E:\\常用URL.txt");
        BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
        response.setHeader("content-type", "application/octet-stream");
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment;filename= 0" + file.getName());
        byte[] b = new byte[1024];
        OutputStream os = response.getOutputStream();
        while((bis.read(b)) != -1) {
            os.write(b);
        }
        os.flush();
        bis.close();
        return "ok";
    }

}
