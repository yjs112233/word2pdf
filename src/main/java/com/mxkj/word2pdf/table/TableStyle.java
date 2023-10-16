package com.mxkj.word2pdf.table;


import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Builder
@Data
@NoArgsConstructor
@AllArgsConstructor
public class TableStyle {

    /**
     *  字号大小
     */
    private Integer fontSize;

    /**
     *  字体
     */
    private String fontFamily;
}
