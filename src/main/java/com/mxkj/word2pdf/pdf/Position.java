package com.mxkj.word2pdf.pdf;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Position {

    /**
     *  横坐标
     */
    private String x;

    /**
     *  纵坐标
     */
    private String y;

    /**
     *  所在页面
     */
    private Integer page;

    /**
     *  指定的标识符
     */
    private String charact;
}
