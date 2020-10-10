package com.zhkusfsc.sqltodoc;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.core.annotation.Order;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.sql.Connection;
import java.sql.ResultSet;
import java.util.*;

/**
 * @author: 史创雄
 * @create: 2020-10-10 14:23
 */

@Component
@Slf4j
@Order(value = 999)
public class InitProject implements CommandLineRunner {

    @Autowired
    JdbcTemplate jdbcTemplate;

    private List<TableInfo> list = new ArrayList<>();
    private String systemName = "测试系统";
    private String userName = "史创雄";
    private String dateTime = "2020-10-10";
    List<String> rowValueList;

    @Override
    public void run(String... args) throws Exception {
        //获取数据库表
        List<Map<String, Object>> ll = jdbcTemplate.queryForList(
                "select table_name,table_comment from information_schema.tables where table_schema='pureshare_base'");
        ll.forEach(t -> {
            try {
                //获取单个表结构信息
                getTableInfo(String.valueOf(t.get("table_name")),String.valueOf(t.get("table_comment")));
            } catch (Exception e) {
                e.printStackTrace();
            }
        });

        //替换字符串
        Map<String,String> replaceMap = new HashMap<>();
        replaceMap.put("{system_name}",systemName);
        replaceMap.put("{user_name}",userName);
        replaceMap.put("{date_time}",dateTime);

        //生成文档
        createDoc(replaceMap);
    }

    private void createDoc(Map<String,String> map) throws IOException {
        String basePath = System.getProperty("user.dir");
        File f = new File(basePath+File.separator+"docx");
        if(!f.exists() || !f.isDirectory()){
            f.mkdirs();
        }
        String destPath = basePath+File.separator+"docx"+File.separator+"test_"+System.currentTimeMillis()+".docx";

        XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(basePath+File.separator+"test.docx"));
        // 替换段落中的指定文字
        Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
        while (itPara.hasNext()) {
            XWPFParagraph paragraph = (XWPFParagraph) itPara.next();
            List<XWPFRun> runs = paragraph.getRuns();
            String temp = "";
            for (int i = 0; i < runs.size(); i++) {
                String oneparaString = runs.get(i).getText(runs.get(i).getTextPosition());
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    if (oneparaString != null && oneparaString.contains(entry.getKey())) {
                        oneparaString = oneparaString.replace(entry.getKey(), entry.getValue());
                    }
                }
                runs.get(i).setText(oneparaString, 0);
            }
        }


        List<String> titleList = new ArrayList<>();
        titleList.add("字段");
        titleList.add("类型");
        titleList.add("是否可空");
        titleList.add("注释");

        //标题单元格样式
        CellStyle titleStyle = new CellStyle();
        titleStyle.setAlignment(ParagraphAlignment.CENTER);
        titleStyle.setBold(true);
        titleStyle.setColor("000000");
        titleStyle.setFontSize(16);
        titleStyle.setVertAlign(XWPFTableCell.XWPFVertAlign.CENTER);
        titleStyle.setFontFamily("宋体");
        titleStyle.setHeight(20);

        //普通单元格样式
        CellStyle commonStyle = new CellStyle();
        commonStyle.setAlignment(ParagraphAlignment.CENTER);
        commonStyle.setBold(false);
        commonStyle.setColor("444444");
        commonStyle.setFontSize(12);
        commonStyle.setVertAlign(XWPFTableCell.XWPFVertAlign.CENTER);
        commonStyle.setFontFamily("宋体");
        commonStyle.setHeight(16);

        document.createParagraph().createRun().addBreak(BreakType.PAGE);//换页

        //表信息
        for(int i=0;i<list.size();i++){
            TableInfo s = list.get(i);
            //表头
            //创建一个段落
            XWPFParagraph para = document.createParagraph();

            //一个XWPFRun代表具有相同属性的一个区域：一段文本
            XWPFRun run = para.createRun();
            run.setBold(true); //加粗
            run.setFontSize(20);
            run.setText(s.getTableName()+"("+s.getTableComment()+")");

            //表格
            XWPFTable table = document.createTable(1, 4);
            table.setWidth(8310);
            setRows(titleList,table.getRow(0),titleStyle);

            //遍历字段
            s.getList().forEach(t -> {
                rowValueList = new ArrayList<>();
                rowValueList.add(t.getColumnName());
                rowValueList.add(t.getTypeName().toLowerCase()+"("+t.getColumnSize()+")");
                rowValueList.add("1".equals(t.getNullable()) ? "是" : "否");
                rowValueList.add(t.getRemarks());
                setRows(rowValueList,table.createRow(),commonStyle);
            });
            if(i < list.size() -1){
                document.createParagraph().createRun().addBreak();//换行
            }
        }

        //文件输出
        FileOutputStream outStream = new FileOutputStream(destPath);
        document.write(outStream);
        outStream.close();
    }

    /**
     * 设置行信息
     * @param valueList
     * @param row
     * @param style
     */
    private void setRows(List<String> valueList, XWPFTableRow row,CellStyle style){
        List<XWPFTableCell> cells = row.getTableCells();
        for(int i=0;i<valueList.size()&&i<cells.size();i++){
            setText(cells.get(i),valueList.get(i),style,i);
        }
    }

    /**
     * 设置表格单元格：文字、样式等
     * @param cell 单元格
     * @param text 文本
     * @param style 样式
     */
    private static void setText(XWPFTableCell cell,String text,CellStyle style,int index){
        if(null == text){
            return;
        }

        //创建段落
        XWPFParagraph p = cell.getParagraphs().get(0);
        p.setAlignment(style.getAlignment());
        cell.setVerticalAlignment(style.getVertAlign());
        cell.setParagraph(p);

        XWPFRun run = null;
        if(cell.getParagraphs().size() > 0 && cell.getParagraphs().get(0).getRuns().size() > 0){
            run = cell.getParagraphs().get(0).getRuns().get(0);
        }else{
            run = p.createRun();
        }
        CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();

        CTTcPr tcpr = cell.getCTTc().addNewTcPr();
        CTTblWidth cellw = tcpr.addNewTcW();
        cellw.setType(STTblWidth.DXA);
        if(index < 3){
            cellw.setW(BigInteger.valueOf(360*5));
        }else{
            cellw.setW(BigInteger.valueOf(360*10));
        }

        //样式设置
        fonts.setAscii(style.getFontFamily());
        fonts.setEastAsia(style.getFontFamily());
        fonts.setHAnsi(style.getFontFamily());
        run.setFontSize(style.getFontSize());
        run.setFontFamily(style.getFontFamily());
        run.setBold(style.isBold());
        run.setColor(style.getColor());
        String[] strs = text.split("\n");
        for(int i=0;i<strs.length;i++){
            if(i > 0){
                run.addBreak();
            }
            run.setText(strs[i],i);
        }
    }

    /**
     * 获取表结构
     * @param tableName
     * @param tableComment
     * @throws Exception
     */
    private void getTableInfo(String tableName,String tableComment) throws Exception {
        List<ColumnInfo> columneList = new ArrayList<>();
        Connection conn = jdbcTemplate.getDataSource().getConnection();
        ResultSet rs = conn.getMetaData().getColumns(null, getSchema(conn),tableName.toUpperCase(), "%");
        while(rs.next()){
            columneList.add(new ColumnInfo(
                    rs.getString("COLUMN_NAME"),//列名
                    rs.getString("DATA_TYPE"),//数据类型
                    rs.getString("TYPE_NAME"),//类型名称
                    rs.getString("COLUMN_SIZE"),//大小限制
                    rs.getString("NULLABLE"),//是否可空
                    rs.getString("REMARKS"))); //备注
        }

        //表信息
        TableInfo tableInfo = new TableInfo();
        tableInfo.setTableName(tableName);
        tableInfo.setTableComment(tableComment);
        tableInfo.setList(columneList);

        list.add(tableInfo);
    }

    /**
     * 判断数据库是否支持
     * @param conn
     * @return
     * @throws Exception
     */
    private String getSchema(Connection conn) throws Exception {
        String schema;
        schema = conn.getMetaData().getUserName();
        if ((schema == null) || (schema.length() == 0)) {
            throw new Exception("ORACLE数据库模式不允许为空");
        }
        return schema.toUpperCase().toString();

    }
}
