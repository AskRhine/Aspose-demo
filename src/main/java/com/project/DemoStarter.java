package com.project;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

/**
 * @Author Ask
 * @Date 2022/2/7
 * @Describe
 */
@SpringBootApplication(exclude = {DataSourceAutoConfiguration.class})
public class DemoStarter {
    public static void main(String[] args) {
        SpringApplication.run(DemoStarter.class, args);
    }
}
