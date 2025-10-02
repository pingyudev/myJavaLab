package com.example.myjavalab.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class HelloController {

    @GetMapping("/hello")
    public String hello() {
        return "Hello world!";
    }

    @GetMapping("/helloYou")
    public String helloYou(@RequestParam("username") String username) {
        return "Hello, " + username;
    }
}