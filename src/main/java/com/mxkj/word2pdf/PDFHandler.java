package com.mxkj.word2pdf;

import com.mxkj.word2pdf.pdf.Position;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class PDFHandler {

    public static List<Position> getPosition(InputStream inputStream, String... characts) throws IOException {
        if (characts == null){
            throw new Word2PdfException("字符参数不能为空");
        }
        for (int i = 0; i < characts.length; i++) {
            if (characts[i] == null || characts[i].isEmpty()){
                throw new Word2PdfException("字符参数不能为空");
            }
            if (characts[i].length() > 1){
                throw new Word2PdfException("只能设置单个字符");
            }
        }

        PDDocument document = PDDocument.load(inputStream);
        List<Position> list = new ArrayList<>();
        // 创建PDFTextStripper对象
        PDFTextStripper stripper = new PDFTextStripper() {
            @Override
            protected void writeString(String text, List<TextPosition> textPositions) throws IOException {
                // 遍历文本位置列表
                for (TextPosition textPosition : textPositions) {
                    // 检查是否为特殊字符
                    for (String charact : characts) {
                        if (textPosition.getUnicode().equals(charact)) {
                            Position position = new Position();
                            position.setX(String.valueOf(textPosition.getX()));
                            position.setY(String.valueOf(textPosition.getY()));
                            position.setPage(this.getCurrentPageNo());
                            position.setCharact(charact);
                            list.add(position);
                        }
                    }
                }
            }
        };

        // 提取文本内容和坐标信息
        stripper.setSortByPosition(true);
        stripper.setStartPage(1);
        stripper.setEndPage(document.getNumberOfPages());
        stripper.getText(document);

        // 关闭文档
        document.close();
        if (list.isEmpty()){
            throw new Word2PdfException("未发现指定字符");
        }
        return list;
    }

    private static void main(String[] args) throws IOException {

    }
}
