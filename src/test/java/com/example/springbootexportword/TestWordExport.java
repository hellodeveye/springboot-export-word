package com.example.springbootexportword;

import cn.afterturn.easypoi.word.WordExportUtil;
import cn.afterturn.easypoi.word.entity.MyXWPFDocument;
import cn.afterturn.easypoi.word.entity.WordImageEntity;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.FileOutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * @author kim.yang
 * @date 2021/11/6 11:02 上午
 */
public class TestWordExport {

    @Test
    public void imageWordExport() {
        Map<String, Object> map = new HashMap<>();
        map.put("test", "Easypoi");
        try {
            Resource resource = new ClassPathResource("templates/风机总损失电量为.docx");
            XWPFDocument document = new MyXWPFDocument(resource.getInputStream());
            WordExportUtil.exportWord07(
                    document, map);
            FileOutputStream fos = new FileOutputStream("/Users/duke/test.docx");
            document.write(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
