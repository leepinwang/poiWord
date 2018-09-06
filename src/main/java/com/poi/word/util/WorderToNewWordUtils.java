package com.poi.word.util;

import java.io.*;
import java.util.*;
import java.util.Map.Entry;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.springframework.util.ResourceUtils;
import org.springframework.util.StringUtils;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

/**
 * 通过word模板生成新的word工具类
 *
 * @author leepinwang
 */
public class WorderToNewWordUtils {

    /**
     * 根据模板生成新word文档
     * 判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
     *
     * @param resourceLocation 模板存放地址
     * @param fileName         文件的名字
     * @param textMap          需要替换的信息集合
     * @param response         HttpServletResponse
     * @return 成功返回true, 失败返回false
     */
    public static boolean changWord(String resourceLocation, String fileName,
                                    Map<String, String> textMap, List<Map<String, String>> tableList, HttpServletResponse response) {

        //模板转换默认成功
        boolean changeFlag = true;
        try {
            //获取docx解析对象
            //springboot获取模版文件
            File fis = ResourceUtils.getFile(resourceLocation);
            InputStream is = new FileInputStream(fis);
            XWPFDocument document = new XWPFDocument(is);
            //解析替换文本段落对象
            WorderToNewWordUtils.changeText(document, textMap);
            //解析替换表格对象
            WorderToNewWordUtils.changeTable(document, textMap, tableList);
            //以文件下载的格式导出word
            downLoadWord(document, resourceLocation, fileName, response);
        } catch (IOException e) {
            e.printStackTrace();
            changeFlag = false;
        }

        return changeFlag;
    }

    /**
     * 下载word文档
     *
     * @param document         docx解析对象
     * @param resourceLocation 模版文档的地址
     * @param fileName         下载文件的名字
     * @param response         HttpServletResponse
     */
    public static void downLoadWord(XWPFDocument document, String resourceLocation, String fileName, HttpServletResponse response) {
        try {
            File fis = ResourceUtils.getFile(resourceLocation);

            response.reset();
            response.setContentType("application/octet-stream");
            response.setContentType("application/x-msdownload");
            fileName = new String(fileName.getBytes(), "ISO8859-1"); //正确,不发生乱码
            //response.addHeader("Content-Disposition", "attachment; filename=\"" + URLEncoder.encode(fileName) + "\"");
            // 特殊符号，生成文件后，无法正确转码回来
            response.addHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
            ByteArrayOutputStream ostream = new ByteArrayOutputStream();
            ServletOutputStream servletOS = response.getOutputStream(); //输出流
            //数据写入到输出流
            document.write(ostream);
            servletOS.write(ostream.toByteArray());
            servletOS.flush();
            servletOS.close();
            ostream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }


    /**
     * 替换段落文本
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     */
    public static void changeText(XWPFDocument document, Map<String, String> textMap) {
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                String temp = "";
                List<Integer> listIndex = new ArrayList();
                for (int i = 0; i < runs.size() - 1; i++) {
                    XWPFRun xwpfRun = runs.get(i);
                    String runText = xwpfRun.toString();
                    if (StringUtils.isEmpty(runText)) {//如果runText是空，那么就跳过本次循环
                        continue;
                    }
                    //存在${ 或者｝
                    if (runText.indexOf("${") != -1 || runText.indexOf("}") != -1) {
                        if (runText.indexOf("${") != -1 && runText.indexOf("}") != -1) {
                            //同时存在，说明是一个完整的${}，在这里就应该先进行替换
                            xwpfRun.setText(changeValue(runText, textMap), 0);
                        } else {
                            listIndex.add(i);
                        }
                    }
                }
                //得到了${ 和}在runs中的位置信息
                for (int i = 0; i < listIndex.size() - 1; i++) {
                    int begin = listIndex.get(i);
                    int end = listIndex.get(++i);
                    String temp1 = "";
                    for (int a = begin; a <= end; a++) {
                        temp1 += runs.get(a).toString();
                        //将run清空
                        runs.get(a).setText("", 0);
                    }
                    runs.get(begin).setText(changeValue(temp1, textMap), 0);
                }
            }
        }
    }


    /**
     * 替换表格对象方法
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     */
    public static void changeTable(XWPFDocument document, Map<String, String> textMap, List<Map<String, String>> tableList) {
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {
            //只处理行数大于等于2的表格，且不循环表头
            XWPFTable table = tables.get(i);
            if (table.getRows().size() > 1) {
                //判断表格是需要替换还是需要插入，判断逻辑有$
                if (checkText(table.getText())) {
                    List<XWPFTableRow> rows = table.getRows();
                    //遍历表格,并替换模板
                    eachTable(rows, textMap);
                }
                //处理动态插入行的部分，此时这个区域无${},这个要根据具体的业务来处理了

                //往没有${}标签的地方动态插入数据
                List<List<String>> list = new ArrayList<>();
                for (int a = 0; a < tableList.size(); a++) {
                    Map<String, String> map = tableList.get(a);
                    List<String> tempList = new ArrayList<>();
                    //第一列的值设置为""
                    tempList.add("");
                    String part = objToStr(map.get("periodDate"));
                    tempList.add(part);
                    String model = objToStr(map.get("school"));
                    tempList.add(model);
                    String num = objToStr(map.get("position"));
                    tempList.add(num);
                    list.add(tempList);
                }
                insertTable(table, list, 4, 7);
            }
        }
    }

    /**
     * 为表格插入数据，行数不够添加新行
     *
     * @param table          需要插入数据的表格
     * @param tableList      插入数据集合
     * @param initRows       要动态插入数据区域的原始行数，
     * @param beginInsertRow 开始插入的行数的地方
     */
    public static void insertTable(XWPFTable table, List<List<String>> tableList, int initRows, int beginInsertRow) {
        //1、先根据数据行数增加行
        int tableListLength = tableList.size();
        //增加行之前
        if (tableListLength > initRows) {
            addOrRemoveRow(table, tableListLength - initRows, beginInsertRow);
        }
        //遍历表格插入数据
        int endRow = beginInsertRow + tableListLength;
        for (int i = beginInsertRow; i < endRow; i++) {
            XWPFTableRow newRow = table.getRow(i);
            List<XWPFTableCell> cells = newRow.getTableCells();
            List<String> arrayValue = tableList.get(i - (beginInsertRow));
            for (int j = 0; j < cells.size(); j++) {
                XWPFTableCell cell = cells.get(j);
                cell.setText(arrayValue.get(j));
            }
        }
        //合并
        mergeCellsVertically(table, 0, beginInsertRow - 1, endRow - 1);
    }


    /**
     * object转String，无法转换返回""
     *
     * @param value
     * @return
     */
    public static String objToStr(Object value) {
        return value != null ? value.toString() : "";
    }

    /**
     * 遍历表格
     *
     * @param rows    表格行对象
     * @param textMap 需要替换的信息集合
     */
    public static void eachTable(List<XWPFTableRow> rows, Map<String, String> textMap) {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if (checkText(cell.getText())) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        String paragraphText = paragraph.getText();
                        //获取当前paragraph，有多个要替换的区域
                        int count = getCharacterCounts(paragraphText, "$");
                        List<XWPFRun> runs = paragraph.getRuns();
                        int runsLength = runs.size();
                        paragraphText = paragraphText.trim();
                        //只有单个$
                        if (count == 1) {
                            String temp = paragraphText;
                            for (XWPFRun run : runs) {
                                run.setText("", 0);
                            }
                            runs.get(0).setText(changeValue(temp, textMap), 0);
                        } else { //count 大于1的时候，即是有多个${}的情况
                            String temp = "";
                            List<Integer> listIndex = new ArrayList();
                            for (int i = 0; i < runs.size() - 1; i++) {
                                XWPFRun xwpfRun = runs.get(i);
                                String runText = xwpfRun.toString();
                                if (StringUtils.isEmpty(runText)) {//如果runText是空，那么就跳过本次循环
                                    continue;
                                }
                                if (runText.indexOf("${") != -1 || runText.indexOf("}") != -1) {//存在${
                                    listIndex.add(i);
                                }
                            }
                            //得到了${ 和}在runs中的位置信息
                            for (int i = 0; i < listIndex.size() - 1; i++) {
                                int begin = listIndex.get(i);
                                int end = listIndex.get(++i);
                                String temp1 = "";
                                for (int a = begin; a <= end; a++) {
                                    temp1 += runs.get(a).toString();
                                    //将run清空
                                    runs.get(a).setText("", 0);
                                }
                                runs.get(begin).setText(changeValue(temp1, textMap), 0);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 统计一个字符在字符串中出现的次数
     */
    public static int getCharacterCounts(String inputs, String singleString) {
        int count = 0;
        for (int i = 0; i <= inputs.length() - 1; i++) {
            String g = inputs.substring(i, i + 1);
            if (g.equals(singleString)) {
                count++;
            }
        }
        return count;
    }

    /**
     * 为表格插入数据，行数不够添加新行
     *
     * @param table     需要插入数据的表格
     * @param tableList 插入数据集合
     */
    public static void insertTable(XWPFTable table, List<String[]> tableList) {
        //创建行,根据需要插入的数据添加新行，不处理表头
        for (int i = 1; i < tableList.size(); i++) {
            XWPFTableRow row = table.createRow();
        }
        //遍历表格插入数据
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 1; i < rows.size(); i++) {
            XWPFTableRow newRow = table.getRow(i);
            List<XWPFTableCell> cells = newRow.getTableCells();
            for (int j = 0; j < cells.size(); j++) {
                XWPFTableCell cell = cells.get(j);
                cell.setText(tableList.get(i - 1)[j]);
            }
        }
    }


    /**
     * 判断文本中时候包含$
     *
     * @param text 文本
     * @return 包含返回true, 不包含返回false
     */
    public static boolean checkText(String text) {
        boolean check = false;
        if (text.indexOf("$") != -1) {
            check = true;
        }
        return check;

    }

    /**
     * 匹配传入信息集合与模板
     *
     * @param value   模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static String changeValue(String value, Map<String, String> textMap) {
        Iterator<Entry<String, String>> it = textMap.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, String> entry = it.next();
            //使用迭代器的remove()方法删除元素
            String key = "${" + entry.getKey() + "}";
            if (value.indexOf(key) != -1) {
                value = entry.getValue();
                //将值从map中删除，提高效率
                it.remove();
                break;//跳出该while循环
            }
        }
        //模板未匹配到区域替换为空
        if (checkText(value)) {
            value = "";
        }
        return value;
    }

    /**
     * 增加或者删除表格的行
     *
     * @param table   被操作的表格对象
     * @param add     要增加的行数，正数表示增加，负数表示删除
     * @param fromRow 从哪一行开始增加增加，如果是增加行，则fromRow是开始增加行的最小值，如果是删除行，则fromRow是开始删除行的最大值
     * @return void
     */
    public static void addOrRemoveRow(XWPFTable table, int add, int fromRow) {
        //XWPFTableRow row = table.getRow(fromRow - 1);
        XWPFTableRow row = table.getRow(fromRow);
        if (add > 0) {
            while (add > 0) {
                copyPro(row, table.insertNewTableRow(fromRow));
                add--;
            }
        } else {
            while (add < 0) {
                table.removeRow(fromRow - 1);
                add++;
            }
        }
    }

    /**
     * 增加行的时候，复制属性
     *
     * @param sourceRow 源行
     * @param targetRow 新增的行
     * @return void
     */
    public static void copyPro(XWPFTableRow sourceRow, XWPFTableRow targetRow) {
        //复制行属性
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        List<XWPFTableCell> cellList = sourceRow.getTableCells();
        if (null == cellList) {
            return;
        }
        //添加列、复制列以及列中段落属性
        XWPFTableCell targetCell = null;
        for (XWPFTableCell sourceCell : cellList) {
            targetCell = targetRow.addNewTableCell();
            //列属性
            targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            //段落属性
            targetCell.getParagraphs().get(0).getCTP().setPPr(sourceCell.getParagraphs().get(0).getCTP().getPPr());
        }
    }


    /**
     * word跨行并单元格
     *
     * @param table   要操作的表格
     * @param col     合并行所在的列
     * @param fromRow 开始合并的行号
     * @param toRow   合并的结束行
     * @return void
     */
    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }


    public void downLoadWord() {
        //模板文件地址
        String inputUrl = "D:\\002.docx";
        //新生产的模板文件
        String outputUrl = "D:\\test.docx";

        Map<String, String> testMap = new HashMap<String, String>();
        testMap.put("name", "小明");
        testMap.put("sex", "男");
        testMap.put("address", "软件园");
        testMap.put("phone", "88888888");
        testMap.put("email", "1369759743@qq.com");

        List<String[]> testList = new ArrayList<String[]>();
        testList.add(new String[]{"1", "1AA", "1BB", "1CC"});
        testList.add(new String[]{"2", "2AA", "2BB", "2CC"});
        testList.add(new String[]{"3", "3AA", "3BB", "3CC"});
        testList.add(new String[]{"4", "4AA", "4BB", "4CC"});

        //WorderToNewWordUtils.changWord(inputUrl, outputUrl, testMap, testList);
    }
}

