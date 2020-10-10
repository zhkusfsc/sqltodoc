package com.zhkusfsc.sqltodoc;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * @author: 史创雄
 * @create: 2020-10-10 14:58
 */

@Data
public class TableInfo {

    private String tableName;
    private String tableComment;
    private List<ColumnInfo> list;

    public TableInfo() {
        list = new ArrayList<>();
    }
}
