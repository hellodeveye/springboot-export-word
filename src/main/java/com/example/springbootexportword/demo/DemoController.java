package com.example.springbootexportword.demo;

import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.word.WordExportUtil;
import cn.afterturn.easypoi.word.entity.MyXWPFDocument;
import com.alibaba.excel.EasyExcel;
import org.apache.commons.io.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Map;

/**
 * @author kim.yang
 * @date 2021/11/5 9:27 下午
 */
@Controller
public class DemoController {


    @GetMapping("/download")
    public void download(HttpServletResponse response) throws Exception {
        String fileName = "大部件巡查模板.xlsx";
        try (ServletOutputStream outputStream = response.getOutputStream()) {
            //获取页面输出流
            Resource resource = new ClassPathResource("templates/大部件巡查模板.xlsx");
            //向输出流写文件
            //写之前设置响应流以附件的形式打开返回值,这样可以保证前边打开文件出错时异常可以返回给前台
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));
            response.setContentType("application/octet-stream");
            //读取文件
            IOUtils.copy(resource.getInputStream(), outputStream);
            outputStream.flush();
        }
    }

    @GetMapping("export")
    public void export(HttpServletResponse response) throws Exception {
        String fileName = "风机总损失电量为.docx";
        Map<String, Object> map = new HashMap<>();
        map.put("test", "Easypoi");
        Resource resource = new ClassPathResource("templates/风机总损失电量为.docx");
        try (ServletOutputStream outputStream = response.getOutputStream();
             XWPFDocument document = new MyXWPFDocument(resource.getInputStream())) {
            WordExportUtil.exportWord07(
                    document, map);
            //向输出流写文件
            //写之前设置响应流以附件的形式打开返回值,这样可以保证前边打开文件出错时异常可以返回给前台
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));
            response.setContentType("application/octet-stream");
            //读取文件
            document.write(outputStream);
            outputStream.flush();
        }
    }

    @GetMapping("exportExcel")
    public void exportExcel(HttpServletResponse response) throws Exception {
        Map<String, Object> map = new HashMap<>();
        map.put("test", "Easypoi");
        Resource resource = new ClassPathResource("templates/EPC常规业务测试表_2G(1).xlsx");
        try (ServletOutputStream outputStream = response.getOutputStream()) {
            EasyExcel.write(outputStream).withTemplate(resource.getInputStream()).sheet().doFill(map);
            //向输出流写文件
            //写之前设置响应流以附件的形式打开返回值,这样可以保证前边打开文件出错时异常可以返回给前台
            // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
            String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("utf-8");
            outputStream.flush();
        }
    }


    @PostMapping("import")
    @ResponseBody
    public Object importExcel(@RequestParam MultipartFile file) throws Exception {
        ImportParams params = new ImportParams();
        params.setTitleRows(1);
        params.setHeadRows(1);
        return ExcelImportUtil.importExcel(
                file.getInputStream(),
                Map.class, params);
    }
}
