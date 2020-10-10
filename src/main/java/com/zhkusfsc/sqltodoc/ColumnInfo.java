package com.zhkusfsc.sqltodoc;

import lombok.Data;

/**
 * @author: 史创雄
 * @create: 2020-10-10 14:58
 */
@Data
public class ColumnInfo {
    private String columnName; // 字段名称
    private String dataType; // 字段类型
    private String typeName; // 字段类型名称
    private Object columnSize; // 值
    private String nullable;
    private String remarks;

    public ColumnInfo(String columnName, String dataType, String typeName, Object columnSize, String nullable, String remarks) {
        this.columnName = columnName;
        this.dataType = dataType;
        this.typeName = typeName;
        this.columnSize = columnSize;
        this.nullable = nullable;
        this.remarks = remarks;
    }

    public ColumnInfo() {
    }
}
