package com.project.table;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

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

    @PostMapping("delCell/{row}/{cell}")
    public void delCell(@PathVariable int row, @PathVariable int cell) throws Exception {
        demoService.delOneCell(row, cell);
    }

    @PostMapping("delCell/{row}")
    public void margeCell(@PathVariable int row) throws Exception {
        demoService.margeOneCell(row);
    }

}
