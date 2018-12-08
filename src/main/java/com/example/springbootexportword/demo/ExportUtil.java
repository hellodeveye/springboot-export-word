package com.example.springbootexportword.demo;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.FileUtils;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.net.URISyntaxException;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * ${DESCRIPTION}
 *
 * @author 杨凯
 * @create 2017-12-25 14:54
 **/
@Slf4j
public class ExportUtil {

	public static String outDocx(String ftlTemplateName, Object dataMap, HttpServletRequest request) throws IOException, URISyntaxException, TemplateException {
		//获取模板
		//创建配置实例
		Configuration configuration = new Configuration();
		//设置编码
		configuration.setDefaultEncoding("UTF-8");
		//设置模板保存路径
		configuration.setClassForTemplateLoading(ExportUtil.class, "/templates/");
		//获取模板
		Template template = configuration.getTemplate(ftlTemplateName);

		//临时文件
		String outTempFtlFilePath = request.getServletContext().getRealPath("/export/emp/docx/temp.ftl");
		File outFtlFile = new File(outTempFtlFilePath);
		//如果输出目标文件夹不存在，则创建
		if (!outFtlFile.getParentFile().exists()) {
			boolean mkdirs = outFtlFile.getParentFile().mkdirs();
			if (!mkdirs) {
				log.error("word导出===》创建文件夹出现异常");
			}
		}

		//填充完数据的临时ftl
		Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFtlFile), "UTF-8"));
		//合并数据
		template.process(dataMap, writer);
		writer.close();
		//输出文件路径
		String outFilePath = "";
		ZipFile zipFile = null;
		InputStream is = null;
		InputStream in = null;
		ZipOutputStream zipOut = null;
		try {
			//空白Word 文件 用于替换新数据
			String fileName = "template.docx";
			Resource resource = new ClassPathResource("static/"+fileName);
			InputStream inputStream = resource.getInputStream();
			String templateFile = request.getServletContext().getRealPath("/export/emp/docx/template/");
			File file = new File(templateFile, fileName);
			if (!file.exists()) {
				try {
					boolean mkdirs = file.getParentFile().mkdirs();
					if (!mkdirs) {
						log.error("word导出===》创建文件夹出现异常");
					}
					boolean newFile = file.createNewFile();
					if (!newFile) {
						log.error("word导出===》创建文件出现异常");
					}
					FileUtils.copyInputStreamToFile(inputStream, file);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			zipFile = new ZipFile(file);
			Enumeration<? extends ZipEntry> zipEntrys = zipFile.entries();
			outFilePath = request.getServletContext().getRealPath("/export/emp/docx/temp.docx");
			zipOut = new ZipOutputStream(new FileOutputStream(outFilePath));
			int len = -1;
			byte[] buffer = new byte[1024];
			while (zipEntrys.hasMoreElements()) {
				ZipEntry next = zipEntrys.nextElement();
				is = zipFile.getInputStream(next);
				//把输入流的文件传到输出流中 如果是word/document.xml由我们输入
				zipOut.putNextEntry(new ZipEntry(next.toString()));
				if ("word/document.xml".equals(next.toString())) {
					in = new FileInputStream(outFtlFile);
					while ((len = in.read(buffer)) != -1) {
						zipOut.write(buffer, 0, len);
					}
				} else {
					while ((len = is.read(buffer)) != -1) {
						zipOut.write(buffer, 0, len);
					}
				}
			}
		} catch (FileNotFoundException e) {
			log.error("word导出===》找不到文件[{}]",e.getMessage());
		} finally {
			try {
				if (in != null) {
					is.close();
				}
				if (is != null) {
					is.close();
				}
				if (zipOut != null) {
					zipOut.close();
				}
				if(zipFile != null){
					zipFile.close();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return outFilePath;
	}
}
