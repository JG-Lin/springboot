package com.demo.springboot.controller;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

//@RestController
public class SampleController {
//    @RequestMapping("/")
    public String sample() {
        return "test";
    }
}
