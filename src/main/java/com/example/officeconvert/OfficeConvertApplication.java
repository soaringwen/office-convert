package com.example.officeconvert;

import com.example.officeconvert.config.ConversionProperties;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;

@SpringBootApplication
@EnableConfigurationProperties(ConversionProperties.class)
public class OfficeConvertApplication {
    public static void main(String[] args) {
        SpringApplication.run(OfficeConvertApplication.class, args);
    }
}
