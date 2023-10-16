package com.mxkj.word2pdf.pdf;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class PDFPosition {

    /**
     *  坐标数组
     */
    private List<Position> positions;

    /**
     *  生成的PDF文档
     */
    private InputStream inputStream;

}
