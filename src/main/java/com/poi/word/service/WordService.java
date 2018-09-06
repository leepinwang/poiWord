package com.poi.word.service;

import com.poi.word.util.WorderToNewWordUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2018/9/4 0004.
 */
@Service(value = "WordService")
public class WordService {

    public void downLoadWord(HttpServletRequest request, HttpServletResponse response) throws Exception {
        //读取word源文件
        FileInputStream fileInputStream = new FileInputStream("d:/mobanFile.docx");
        // POIFSFileSystem pfs = new POIFSFileSystem(fileInputStream);
        XWPFDocument document = new XWPFDocument(fileInputStream);
        //获取所有表格
        List<XWPFTable> tables = document.getTables();
        //这里简单取第一个表格
        XWPFTable table = tables.get(0);
        //表格的插入行有两种方式，这里使用addNewRowBetween，因为这样会保留表格的样式，就像我们在word文档的表格中插入行一样。注意这里不要使用insertNewTableRow方法插入新行，这样插入的新行没有样式，很难看
        //table.addNewRowBetween(0, 1); 源码没有实现

        //模拟从前端传过来的数据
        //1.人员的基本信息
        Map<String, String> params = new HashMap<String, String>();
        params.put("name1", "刘珩鑫");
        params.put("name", "刘珩鑫");
        params.put("sex", "男");
        params.put("birthday", "1999-09-03");

        params.put("IDCard", "860000458292382937");
        params.put("national", "汉族");
        params.put("landscape", "党员");

        params.put("married", "未婚");
        params.put("healthy", "健康");
        params.put("height", "172cm");

        params.put("address", "西安市碑林区南二环中路185号");
        params.put("zhuanYe", "社会工作");
        params.put("xueLi", "本科");

        params.put("school", "清华大学");
        params.put("graduationDate", "2018-07-01");
        params.put("jishuzhicheng", "高级工程师");

        params.put("wordUnit", "腾讯股份有限公司");
        params.put("beginWorkDate", "2018-09-01");
        params.put("nowWork", "java高级开发工程师");

        params.put("address1", "广州市天河区沙河顶地铁站");
        params.put("post", "5201025");
        params.put("phone", "1369756789");
        params.put("eMail", "1353379845@qq.com");

        //2、动态新增行的数据
        List<Map<String, String>> tableList = new ArrayList<>();
        Map<String, String> map1 = new HashMap<>();
        map1.put("periodDate", "2005/09/01-2007/07/01");
        map1.put("school", "周皮小学");
        map1.put("position", "班长");

        Map<String, String> map2 = new HashMap<>();
        map2.put("periodDate", "2006/09/01-2008/07/01");
        map2.put("school", "平稳小学");
        map2.put("position", "班长");

        Map<String, String> map3 = new HashMap<>();
        map3.put("periodDate", "2008/09/01-2010/07/01");
        map3.put("school", "双桥一中");
        map3.put("position", "班长");

        Map<String, String> map4 = new HashMap<>();
        map4.put("periodDate", "2009/09/01-2012/07/01");
        map4.put("school", "武鸣高中");
        map4.put("position", "班长");

        Map<String, String> map5 = new HashMap<>();
        map5.put("periodDate", "2012/09/01-2015/07/01");
        map5.put("school", "清华大学");
        map5.put("position", "学生会会长");
        tableList.add(map1);
        tableList.add(map2);
        tableList.add(map3);
        tableList.add(map4);
        tableList.add(map5);


        List<String[]> testList = new ArrayList<>();
        //下载文件的名字
        String fileName = "result.docx";
        //模版文件的路径
        /**
         * 注意：模版中的占位符号，要一次性的写完，例如${name},要从左到右一个字符一个字符的敲，不能拷贝，比如先${},然后再name,
         * 这样子操作，poi在读取${name} 的时候，会被分成${、name、｝三个部分，如果占位符过长，poi读取的时候，也会分成三个部分，
         * 可能是poi的读取word模板的机制不够完善，没有读取源码，原因不明，以上也是通过百度得到的信息，附上几个博客地址：
         * http://www.cnblogs.com/hzw-hym/p/4586311.html
         *http://www.cnblogs.com/qingruihappy/p/8443403.html
         */
        String resourceLocation = "classpath:static/template/mobanFile.docx";
        WorderToNewWordUtils.changWord(resourceLocation, fileName, params, tableList, response);
    }
}
