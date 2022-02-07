package com.project;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @Author Ask
 * @Date 2022/2/7
 * @Describe
 */
@RestController
@RequestMapping
public class DemoController {
    @Autowired
    private DemoService demoService;

    @GetMapping("/readTable")
    public void readTable() throws Exception {
        demoService.readTable();
    }

    @PostMapping("/createTable")
    public void createTable() throws Exception {
        demoService.createTable();
    }


    @PostMapping("/altTable")
    public void altTable() throws Exception {
        demoService.altTable(demoService.readTable());
    }
}
