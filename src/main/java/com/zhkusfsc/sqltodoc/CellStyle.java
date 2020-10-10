package com.zhkusfsc.sqltodoc;

import lombok.Data;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

/**
 * @author: 史创雄
 * @create: 2020-08-27 10:32
 */

@Data
public class CellStyle {
    private ParagraphAlignment alignment;//水平位置
    private XWPFTableCell.XWPFVertAlign vertAlign;//垂直位置
    private int fontSize;//字号
    private String fontFamily;//字体
    private String color;//颜色
    private boolean isBold;//加粗
    private int height;//行高


}
