package com.example.demoapachepoi;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;

@SpringBootApplication
@EnableScheduling
public class DemoApachePoiApplication {

	public static void main(String[] args) {
		SpringApplication.run(DemoApachePoiApplication.class, args);
	}

}
