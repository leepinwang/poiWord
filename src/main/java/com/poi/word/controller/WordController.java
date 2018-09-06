package com.poi.word.controller;

;
import com.poi.word.service.WordService;
import com.poi.word.util.WordUtils;
import com.poi.word.util.WorderToNewWordUtils;
import org.apache.poi.wp.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Controller;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 导出word Demo
 * 本功能无页面：直接在浏览器中输入：http://localhost:8080/wordController/downLoadWord   即可下载word文档
 */
@Scope("prototype")
@Controller
@RequestMapping("/wordController")
public class WordController {
    //  http://localhost:8080/wordController/downLoadWord
    @Autowired
    private WordService wordService;

    @RequestMapping("/downLoadWord")
    public void downLoadWord(HttpServletRequest request, HttpServletResponse response) throws Exception {
        wordService.downLoadWord(request, response);
    }
}
