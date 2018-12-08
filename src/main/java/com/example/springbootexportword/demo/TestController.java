package com.example.springbootexportword.demo;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;

/**
 * ${DESCRIPTION}
 *
 * @author 杨凯
 * @create 2017-12-25 15:47
 **/
@Controller
public class TestController {


	/**
	 * 测试
	 * @return
	 */
	@GetMapping
	@ResponseBody
	public String index(){
		return "hello Springboot";
	}
}
