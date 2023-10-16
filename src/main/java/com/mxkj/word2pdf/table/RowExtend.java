package com.mxkj.word2pdf.table;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class RowExtend<T> {

    private XWPFTable table;

    private int pos;

    private int prepare;

    private List<T> list;
}
