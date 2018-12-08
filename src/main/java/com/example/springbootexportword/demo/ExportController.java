package com.example.springbootexportword.demo;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * ${DESCRIPTION}
 *
 * @author 杨凯
 * @create 2017-12-25 15:54
 **/
@Slf4j
@Controller
@RequestMapping("/export")
public class ExportController {

	@GetMapping(value = "/word", produces = "application/octet-stream")
	public void word(HttpServletRequest request, HttpServletResponse response) {
		InputStream fis = null;
		OutputStream os = null;
		try {
			//文件生成路径
			String outFilePath = "";
			//文件名称
			String fileName = "test.docx";
			Map<String,String> map = new HashMap<>();
			map.put("name","张三");
			map.put("sex","男");
			map.put("age","22");


			//判断试卷类型
			outFilePath = ExportUtil.outDocx("test.ftl", map, request);
			fis = new BufferedInputStream(new FileInputStream(outFilePath));
			byte[] buffer = new byte[fis.available()];
			fis.read(buffer);
			// 清空response
			response.reset();
			response.setContentType("application/octet-stream");
			response.addHeader("Content-Disposition", "attachment; filename=" + new String(fileName.getBytes("GB2312"), "ISO-8859-1"));
			os = new BufferedOutputStream(response.getOutputStream());
			os.write(buffer);
		} catch (Exception e) {
			log.error("word导出===》未知异常[{}]", e.getMessage());
		} finally {
			try {
				if (os != null) {
					os.close();
				}
				if (fis != null) {
					fis.close();
				}
			} catch (IOException e) {
				log.error("word导出===》流关闭异常[{}]", e.getMessage());
			}
		}
	}

}
