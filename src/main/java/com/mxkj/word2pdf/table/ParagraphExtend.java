package com.mxkj.word2pdf.table;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ParagraphExtend<T> {

    private XWPFTable table;

    private int rowPos;

    private int colPos;

    private List<T> list;
}
