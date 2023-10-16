package com.mxkj.word2pdf;

import com.aspose.words.*;
import com.aspose.words.Document;
import com.mxkj.word2pdf.pdf.PDFPosition;
import com.mxkj.word2pdf.pdf.Position;
import com.mxkj.word2pdf.table.ParagraphExtend;
import com.mxkj.word2pdf.table.RowExtend;
import com.mxkj.word2pdf.table.TableStyle;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.*;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Word2PdfBuilder {


    private static final String FONT_FAMILY = "宋体";
    /**
     *  aspose.word Document
     */
    private Document document;

    static {
        FontSourceBase[] bases = new FontSourceBase[2];
        FontSourceBase fontSourceBase0 = build("/static/font/simsun.ttc");
        bases[0] = fontSourceBase0;
        FontSourceBase fontSourceBase1 = build("/static/font/Arial Unicode MS.ttf");
        bases[1] = fontSourceBase1;
        FontSettings.setFontsSources(bases);
    }

    private static FontSourceBase build(String path){
        InputStream inputStream = Word2PdfBuilder.class.getResourceAsStream(path);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            IOUtils.copy(inputStream, outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return new MemoryFontSource(outputStream.toByteArray());
    }

    /**
     *  初始化word模板
     * @param file word模板文件
     * @return
     */
    public static Word2PdfBuilder newInstance(File file) {
        Word2PdfBuilder builder = new Word2PdfBuilder();
        try (FileInputStream inputStream = new FileInputStream(file)){
            builder.licenseCertificate();
            builder.document = new Document(inputStream);
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }
        return builder;
    }

    /**
     *  初始化word模板
     * @param inputStream 初始化word模板文件流
     * @return
     */
    public static Word2PdfBuilder newInstance(InputStream inputStream) {
        Word2PdfBuilder builder = new Word2PdfBuilder();
        try {
            builder.licenseCertificate();
            builder.document = new Document(inputStream);
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return builder;
    }

    /**
     *  替换word模板中变量
     *  eg： 传入的obj对象中存在age字段，specifiedName值为test。则将替换word模板中的{test.age}，其他成员字段同理。
     * @param obj 对象
     * @param specifiedName 指定对象变量名
     */
    public Word2PdfBuilder build(Object obj, String specifiedName){
        try {
            if (obj instanceof List){
                List<Object> list = (List<Object>) obj;
                build(list, specifiedName);
            }else {
                build(obj, specifiedName, null);
            }
        } catch (Exception e) {
            throw new Word2PdfException(e.getMessage());
        }
        return this;
    }

    /**
     *  替换word模板中变量
     *  eg： 如果传入的obj对象中存在age字段，类名称为Student，则将替换word模板中的{student.age}，其他成员字段同理。
     * @param obj 对象
     */
    public Word2PdfBuilder build(Object obj){
        try {
            if (obj instanceof List){
                List<Object> list = (List<Object>) obj;
                build(list);
            }else {
                build(obj, null, null);
            }
        } catch (Exception e) {
            throw new Word2PdfException(e.getMessage());
        }
        return this;
    }

    /**
     * 替换word模板中变量
     * eg: 如果传入的数组长度为2，obj对象中存在age字段，specifiedName值为test。
     *     则将替换word模板中的{test[0].age}、{test[1].age}，其他成员字段同理。
     * @param objs 对象数组
     * @param specifiedName 指定对象变量名
     */
    private Word2PdfBuilder build(List<Object> objs, String specifiedName){
        for (int i = 0; i < objs.size(); i++) {
            try {
                Object obj = objs.get(i);
                build(obj, specifiedName, i);
            } catch (Exception e) {
                throw new Word2PdfException(e.getMessage());
            }
        }
        return this;
    }

    /**
     * 替换word模板中变量
     * eg: 如果传入的数组长度为2，obj对象中存在age字段，类名称为Student，
     *     则将替换word模板中的{student[0].age}、{student[1].age}，其他成员字段同理。
     * @param objs 对象数组
     */
    private Word2PdfBuilder build(List<Object> objs){
        for (int i = 0; i < objs.size(); i++) {
            try {
                Object obj = objs.get(i);
                build(obj, null, i);
            } catch (Exception e) {
                throw new Word2PdfException(e.getMessage());
            }
        }
        return this;
    }


    /**
     *
     * @param list 数组集合
     * @param specifiedName 指定的变量名
     * @param minRowCount 最小的行数，如果list含有3个对象，minRowCount等于5， 那么将会有2行的空白行
     * @param <T>
     * @return
     */
    public <T> Word2PdfBuilder buildRowList(List<T> list, String specifiedName,  int minRowCount){
        if (list == null || list.isEmpty()){
            return this;
        }
        Class clazz = list.get(0).getClass();
        String varableNameLike = String.format("{%s[i].", specifiedName);

        ByteArrayInputStream inputStream = toInputStream();
        try {
            // 抓取需要扩展所有table中的变量行
            List<RowExtend> extendList = new ArrayList<>();
            XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
            for (XWPFTable table : xwpfDocument.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    boolean isPointPos = false;
                    for (XWPFTableCell tableCell : row.getTableCells()) {
                        if (tableCell.getText() != null && tableCell.getText().contains(varableNameLike)){
                            isPointPos = true;
                            break;
                        }
                    }
                    if (isPointPos){
                        int pos = table.getRows().indexOf(row);
                        extendList.add(new RowExtend<>(table, pos, minRowCount, list));
                    }
                }
            }
            // 处理
            for (RowExtend rowExtend : extendList) {
                // 拷贝变量行
                XWPFTableRow sourceRow = rowExtend.getTable().getRow(rowExtend.getPos());
                XmlObject xmlObject = sourceRow.getCtRow().copy();
                // 新增
                int max = Math.max(list.size(), rowExtend.getPrepare());
                for (int i = 1; i <= max; i++) {
                    XWPFTableRow newRow = rowExtend.getTable().insertNewTableRow(rowExtend.getPos() + i);
                    newRow.getCtRow().set(xmlObject);
                }
                // 移除变量行

                rowExtend.getTable().removeRow(rowExtend.getPos());
            }
            // 刷新表格
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            xwpfDocument.write(outputStream);
            XWPFDocument newXwpfDocument = new XWPFDocument(new ByteArrayInputStream(outputStream.toByteArray()));

            for (XWPFTable table : newXwpfDocument.getTables()) {
                List<XWPFTableRow> varableRows = new ArrayList<>();
                // 列表变量所在的行
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell tableCell : row.getTableCells()) {
                        if (tableCell.getText() != null && tableCell.getText().contains(varableNameLike)){
                            varableRows.add(row);
                            break;
                        }
                    }
                }
                // 为变量行赋值
                for (int i = 0; varableRows.size() >= list.size() && i < list.size(); i++) {
                    T obj = list.get(i);
                    XWPFTableRow row = varableRows.get(i);
                    // 抓取变量赋值顺序
                    String var = null;
                    for (XWPFTableCell tableCell : row.getTableCells()) {
                        Field field = null;
                        for (Field declaredField : clazz.getDeclaredFields()) {
                            String variable = String.format("%s[i].%s", specifiedName , declaredField.getName());
                            if (tableCell.getText().contains(variable)){
                                field = declaredField;
                                var = tableCell.getText();
                                break;
                            }
                        }
                        // 赋值
                        tableCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                        XWPFParagraph paragraph = tableCell.getParagraphs().get(0);
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        List<XWPFRun> runs = paragraph.getRuns();
                        // 清空原有变量值
                        if (!runs.isEmpty()){
                            for (XWPFRun run : runs) {
                                run.setText("", 0);
                            }
                        }
                        XWPFRun run = runs.isEmpty() ? paragraph.createRun() : runs.get(0);
                        if (field == null){
                            run.setText("", 0);
                            run.setFontFamily(FONT_FAMILY);
                        }else {
                            field.setAccessible(true);
                            Object value = field.get(obj);
                            if (value instanceof List){
                                List<?> ls = (List<?>) value;
                                for (int j = 0; j < ls.size(); j++) {
                                    String variable = String.format("%s[i].%s[%s]", specifiedName , field.getName(), j);
                                    if (var.contains(variable)){
                                        Object o = ls.get(j);
                                        if (isDirectConvert(o)){
                                            run.setText(converObject(ls.get(j)), 0);
                                            run.setFontFamily(FONT_FAMILY);
                                        }else {
                                            Class<?> clacc = o.getClass();
                                            for (Field declaredField : clacc.getDeclaredFields()) {
                                                declaredField.setAccessible(true);
                                                variable = String.format("%s[i].%s[%s].%s", specifiedName , field.getName(), j, declaredField.getName());
                                                if (var.contains(variable)){
                                                    run.setText(converObject(declaredField.get(o)), 0);
                                                    run.setFontFamily(FONT_FAMILY);
                                                    break;
                                                }
                                            }
                                        }
                                        break;
                                    }
                                }
                            }else if(isDirectConvert(value)){
                                run.setText(converObject(value), 0);
                                run.setFontFamily(FONT_FAMILY);
                            }
                        }
                    }
                }
            }
            outputStream.reset();
            newXwpfDocument.write(outputStream);
            document = new Document(new ByteArrayInputStream(outputStream.toByteArray()));
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }
        return this;
    }

    /**
     *  基本类型
     * @param value
     * @return
     */
    private boolean isDirectConvert(Object value){
        return value instanceof Number || value instanceof String || value instanceof Date;
    }

    /**
     * 绑定数组
     * 变量名将默认为类名的小驼峰
     * @param list 数组集合
     * @param minRowCount 最小的行数，如果list含有3个对象，minRowCount等于5， 那么将会有2行的空白行
     * @param <T>
     * @return
     */
    public <T> Word2PdfBuilder buildRowList(List<T> list, int minRowCount){
        if (list == null || list.isEmpty()){
            return this;
        }
        Class clazz = list.get(0).getClass();
        String className = upperToLower(clazz.getSimpleName());
        return buildRowList(list, className, minRowCount);
    }


    /**
     *  某个单元格内的段落循环
     * @param list 集合
     * @param specifiedName 指定变量名
     * @param <T> 类型
     * @return
     */
    public <T> Word2PdfBuilder buildParagraphList(List<T> list, String specifiedName){
        if (list == null || list.isEmpty()){
            return this;
        }
        Class clazz = list.get(0).getClass();
        String varableNameLike = String.format("{%s[i].", specifiedName);

        try {
            ByteArrayInputStream inputStream = toInputStream();
            // 抓取需要扩展所有table中的变量行
            XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
            ParagraphExtend paragraphExtend = getParapragh(xwpfDocument, varableNameLike,list);
            if (paragraphExtend == null){
                return this;
            }
            // 处理
            XWPFTableRow sourceRow = paragraphExtend.getTable().getRow(paragraphExtend.getRowPos());
            XWPFTableCell tableCell = sourceRow.getTableCells().get(paragraphExtend.getColPos());
            List<XWPFParagraph> paramParagraphList = new ArrayList<>(tableCell.getParagraphs());
            if (!paramParagraphList.isEmpty()){
                for (int i = 0; i < list.size() - 1; i++) {
                    // 整段落复制
                    for (XWPFParagraph paragraph : paramParagraphList) {
                        XWPFParagraph newParagraph = tableCell.addParagraph();
                        newParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());
                        for (XWPFRun run : paragraph.getRuns()) {
                            XWPFRun newRun = newParagraph.createRun();
                            newRun.getCTR().set(run.getCTR().copy());
                        }
                    }
                }
            }
            // 刷新表格
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            xwpfDocument.write(outputStream);
            XWPFDocument newXwpfDocument = new XWPFDocument(new ByteArrayInputStream(outputStream.toByteArray()));
            // 替换
            ParagraphExtend newPara = getParapragh(newXwpfDocument, varableNameLike, list);
            XWPFTableCell newTableCell = newPara.getTable().getRow(newPara.getRowPos()).getTableCells().get(newPara.getColPos());
            for (int i = 0; i < list.size() ; i++) {
                // 对象
                T obj = list.get(i);
                XWPFParagraph paragraph = newTableCell.getParagraphs().get(i);
                // 整合成规则段落
                List<XWPFRun> mergeRuns = mergeRuns(paragraph);
                for (XWPFRun run : mergeRuns) {
                    // 文本
                    String content = run.getText(0);
                    if (!content.trim().equals("")){
                        for (Field declaredField : clazz.getDeclaredFields()) {
                            // 属性
                            String variable = String.format("{%s[i].%s}", specifiedName , declaredField.getName());
                            if (content.trim().equals(variable)){
                                declaredField.setAccessible(true);
                                run.setText(converObject(declaredField.get(obj)), 0);
                                break ;
                            }
                        }
                    }
                }
            }

            outputStream.reset();
            newXwpfDocument.write(outputStream);
            document = new Document(new ByteArrayInputStream(outputStream.toByteArray()));
        }catch (Exception e){
            e.printStackTrace();
        }
        return this;
    }

    private <T> ParagraphExtend getParapragh(XWPFDocument xwpfDocument, String varableNameLike, List<T> list){
        ParagraphExtend paragraphExtend = null;
        one: for (XWPFTable table : xwpfDocument.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell tableCell : row.getTableCells()) {
                    if (tableCell.getText() != null && tableCell.getText().contains(varableNameLike)){
                        int rowPos = table.getRows().indexOf(row);
                        int colPos = row.getTableCells().indexOf(tableCell);
                        paragraphExtend = new ParagraphExtend(table, rowPos,colPos, list);
                        break one;
                    }
                }
            }
        }
        return paragraphExtend;
    }

    /**
     *  合并不规则的Run
     * @param paragraph
     * @return
     */
    private static List<XWPFRun> mergeRuns(XWPFParagraph paragraph){
        String content = paragraph.getText();
        String regex = "\\{([^}]+)\\}";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(content);
        List<String> matches = new ArrayList<>();
        while (matcher.find()) {
            matches.add(matcher.group());
        }
        String[] splis = content.replaceAll(regex, "A#A").split("A#A");
        List<String> merges = new ArrayList<>();
        int mt = 0;
        int st = 0;
        for (int i = 0; i < matches.size() + splis.length; i++) {
            if (i % 2 == 0 && splis.length > 0){
                merges.add(splis[st++]);
            }else {
                merges.add(matches.get(mt++));
            }
        }

        // 合并
        List<XWPFRun> list = paragraph.getRuns();
        for (int i = 0; i < list.size(); i++) {
            if (i <= merges.size() - 1){
                list.get(i).setText(merges.get(i), 0);
            }else {
                list.get(i).setText("", 0);
            }
        }
        return list;
    }

    public static void copyRowColumns(XWPFTableRow sourceRow, XWPFTableRow targetRow) {
        // 获取源行中的单元格列表
        List<XWPFTableCell> sourceCells = sourceRow.getTableCells();

        // 遍历源行的单元格
        for (XWPFTableCell sourceCell : sourceCells) {
            // 创建目标单元格
            XWPFTableCell targetCell = targetRow.createCell();

            for (XWPFParagraph sourceParagraph : sourceCell.getParagraphs()) {
                XWPFParagraph targetParagraph = targetCell.addParagraph();
                for (XWPFRun sourceRun : sourceParagraph.getRuns()) {
                    XWPFRun targetRun = targetParagraph.createRun();
                    targetRun.setFontSize(sourceRun.getFontSize());
                    targetRun.setFontFamily(sourceRun.getFontFamily());
                }
            }
        }
    }

    public static void copyColumnMerge(XWPFTableRow sourceRow, XWPFTableRow targetRow) {
        // 获取源行中的单元格列表
        List<XWPFTableCell> sourceCells = sourceRow.getTableCells();

        // 遍历源行的单元格
        for (int i = 0; i < sourceCells.size(); i++) {
            XWPFTableCell sourceCell = sourceCells.get(i);
            XWPFTableCell targetCell = targetRow.getCell(i);

            // 获取源单元格的列合并情况
            boolean isMerged = sourceCell.getCTTc().isSetTcPr() && sourceCell.getCTTc().getTcPr().isSetHMerge();
            if (isMerged) {
                // 获取源单元格的合并信息
                STMerge.Enum mergeType = sourceCell.getCTTc().getTcPr().getHMerge().getVal();
                if (mergeType == STMerge.RESTART) {
                    // 设置目标单元格的合并情况
                    targetCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
                } else if (mergeType == STMerge.CONTINUE) {
                    // 设置目标单元格的合并情况
                    targetCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
                }
            }
        }
    }

    /**
     *  替换第一张图片
     * @param inputStream
     * @return
     */
    public Word2PdfBuilder replaceImage(InputStream inputStream) {
        replaceImage(0, inputStream);
        return this;
    }

    /**
     *  替换指定位置图片
     * @param imageIndex 图片从上到下的下标位置，从0开始
     * @param inputStream 替换的图片流
     * @return
     */
    public Word2PdfBuilder replaceImage(int imageIndex, InputStream inputStream) {
        NodeCollection shapes = document.getChildNodes(NodeType.SHAPE,true);
        if (shapes.getCount() == 0){
            throw new Word2PdfException("word模板中未发现任何图片");
        }
        if (imageIndex > shapes.getCount() - 1){
            throw new Word2PdfException(String.format("下标超限: %s，word模板中最大下标为:%s", imageIndex, shapes.getCount() - 1));
        }
        Shape shape = (Shape) shapes.get(imageIndex);
        try {
            ImageData imageData = shape.getImageData();
            imageData.setImage(inputStream);
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }finally {
            if (inputStream != null){
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return this;
    }


    /**
     *  所有table格式调整
     *  默认宋体8号
     * @return
     * @throws Exception
     */
    public Word2PdfBuilder defaultTableStyle(){
        TableStyle tableStyle = TableStyle.builder()
                .fontSize(8)
                .fontFamily(FONT_FAMILY)
                .build();
        return tableStyle(tableStyle);
    }

    /**
     *  所有table格式调整
     * @param tableStyle
     * @return
     * @throws Exception
     */
    public Word2PdfBuilder tableStyle(TableStyle tableStyle){
        ByteArrayInputStream inputStream = toInputStream();
        try {
            XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
            for (XWPFTable table : xwpfDocument.getTables()) {
                table(table, tableStyle);
            }
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            xwpfDocument.write(outputStream);
            document = new Document(new ByteArrayInputStream(outputStream.toByteArray()));
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }
        return this;
    }


    /**
     *
     *  指定下标的table格式调整
     *  @param tableIndex word模板中的指定table，下标从0开始
     * @param tableStyle table格式
     * @return
     * @throws Exception
     */
    public Word2PdfBuilder tableStyle(int tableIndex, TableStyle tableStyle){
        ByteArrayInputStream inputStream = toInputStream();
        try {
            XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
            for (int i = 0; i < xwpfDocument.getTables().size(); i++) {
                if (i == tableIndex){
                    XWPFTable table = xwpfDocument.getTables().get(i);
                    table(table, tableStyle);
                }
            }
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            xwpfDocument.write(outputStream);
            document = new Document(new ByteArrayInputStream(outputStream.toByteArray()));
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }
        return this;
    }

    /**
     *  为table配置style
     * @param table
     * @param tableStyle
     */
    private void table(XWPFTable table, TableStyle tableStyle){
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell tableCell : row.getTableCells()) {
                for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                    for (XWPFRun run : paragraph.getRuns()) {
                        if (tableStyle.getFontFamily() != null){
                            run.setFontFamily(tableStyle.getFontFamily());
                        }
                        if (tableStyle.getFontSize() != null){
                            run.setFontSize(tableStyle.getFontSize());
                        }
                    }
                }
            }
        }
    }

    /**
     * 替换word模板中变量
     * eg: 如果传入的Map中包含age和name字段，
     *     则将替换word模板中的{age}、{name}，其他成员字段同理。
     * @param map 对象map
     */
    public Word2PdfBuilder buildMap(Map<String, Object> map) throws Exception {
        if (map == null){
            return this;
        }
        return buildMap(map, null);
    }

    /**
     * 替换word模板中变量
     * eg: 如果传入的Map中包含age和name字段，指定名称为abc.
     *     则将替换word模板中的{abc.age}、{abc.name}，其他成员字段同理。
     *     如果指定名称传入null值，则类似@See buildMap(Map<String, Object> map)
     * @param specifiedName 指定前缀
     * @param map 对象map
     */
    public Word2PdfBuilder buildMap(Map<String, Object> map, String specifiedName) throws Exception {
        if (map == null){
            return this;
        }
        String clazzName = specifiedName == null ? "" : String.format("%s.", specifiedName);
        for (Map.Entry<String, Object> entry : map.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();
            if (key != null && value != null){
                String varable = String.format("{%s%s}",clazzName, key);
                document.getRange().replace(varable, converObject(value), false, false);
            }
        }
        return this;
    }

    /***
     * 将{obj.name} 映射为obj对象name成员的值
     * @param obj 传入对象
     * @param specifiedName 是否指定对象变量名，如果指定，则采用指定变量名。
     * @param serial 是否指定下标序号，如果指定，则采用序号。
     * @throws Exception
     */
    private void build(Object obj, String specifiedName, Integer serial) throws Exception {
        if (obj == null){
            return;
        }
        Class clazz = obj.getClass();
        if (obj instanceof List){
            List list = (List) obj;
            if (list.isEmpty()){
                return;
            }
            clazz = list.get(0).getClass();
        }

        String className = upperToLower(clazz.getSimpleName());
        String text = document.getText();
        for (Field declaredField : clazz.getDeclaredFields()) {
            declaredField.setAccessible(true);
            specifiedName = specifiedName == null ? className : specifiedName;
            String variableName = String.format("%s.%s", specifiedName, declaredField.getName());
            if (serial != null){
                variableName = String.format("%s[%s].%s", specifiedName,serial, declaredField.getName());
            }
            String pointName = String.format("{%s}", variableName);
            if (text.contains(variableName)){
                Object value = declaredField.get(obj);
                if (isDirectConvert(value)){
                    document.getRange().replace(pointName, converObject(value), false, false);
                }
                if (value instanceof List){
                    List<?> list = (List<?>) value;
                    for (int i = 0; i < list.size(); i++) {
                        Object object = list.get(i);
                        if (object instanceof String){
                            pointName = String.format("{%s[%s].%s[%s]}", specifiedName, serial, declaredField.getName(), i);
                            document.getRange().replace(pointName, converObject(object), false, false);
                        }else {
                            Class clacc = object.getClass();
                            for (Field field : clacc.getDeclaredFields()) {
                                field.setAccessible(true);
                                Object o = field.get(object);
                                pointName = String.format("{%s[%s].%s[%s].%s}", specifiedName, serial, declaredField.getName(), i, field.getName());
                                document.getRange().replace(pointName, converObject(o), false, false);
                            }
                        }
                    }
                }
            }
        }
    }

    private String converObject(Object value){
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String stringValue = value == null ? "" : value instanceof Date ? dateFormat.format(value) : String.valueOf(value);
        return stringValue;
    }

    /**
     *  大驼峰转小驼峰
     * @param className
     * @return
     */
    private static String upperToLower(String className){
        char[] chars = className.toCharArray();
        char n = Character.toLowerCase(chars[0]);
        return n + className.substring(1);
    }


    /**
     *  输出为docx的word文档
     * @return
     */
    private ByteArrayInputStream toInputStream(){
        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            document.save(outputStream, SaveFormat.DOCX);
            return new ByteArrayInputStream(outputStream.toByteArray());
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }
    }


    private void missedBlank() throws Exception {
        // {}之间的参数且包含{}的正则表达式
        Pattern pattern = Pattern.compile("\\{([^}]+)\\}");
        document.getRange().replace(pattern,"  ");
    }

    /**
     *  输出为docx的word文档
     * @return
     */
    public ByteArrayInputStream toDocx(){
        try {
            missedBlank();
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            document.save(outputStream, SaveFormat.DOCX);
            return new ByteArrayInputStream(outputStream.toByteArray());
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }
    }

    /**
     *  输出为docx的word文档, 保存到file中
     * @return
     */
    public void toDocx(File file){
        saveToFile(file, SaveFormat.DOCX);
    }

    /**
     *  输出为pdf文档
     * @return
     */
    public ByteArrayInputStream toPDF(){
        try {
            missedBlank();
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            document.save(outputStream, SaveFormat.PDF);
            return new ByteArrayInputStream(outputStream.toByteArray());
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }
    }

    /**
     *  输出带指定字符坐标的PDF文档
     * @param chars
     * @return
     * @throws Exception
     */
    public PDFPosition toPDFWithPosition(String... chars) throws Exception {
        InputStream pdfInputStream = toPDF();
        List<Position> positions = PDFHandler.getPosition(pdfInputStream, chars);
        document = document.deepClone();
        for (String aChar : chars) {
            document.getRange().replace(aChar, "", false, false);
        }
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        document.save(outputStream, SaveFormat.PDF);
        return new PDFPosition(positions, new ByteArrayInputStream(outputStream.toByteArray()));
    }

    public void toPDF(File file){
        saveToFile(file, SaveFormat.PDF);
    }

    private void saveToFile(File file, int saveFormat){
        FileOutputStream outputStream = null;
        try {
            if (!file.exists()){
                if (!file.createNewFile()){
                    throw new Word2PdfException("文件创建失败，请检查访问权限");
                }
            }
            missedBlank();
            outputStream = new FileOutputStream(file);
            document.save(outputStream, saveFormat);
        }catch (Exception e){
            throw new Word2PdfException(e.getMessage());
        }finally {
            if (outputStream != null){
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     *  license认证，去除水印
     * @throws Exception
     */
    private void licenseCertificate() throws Exception {
        InputStream is = getClass().getResourceAsStream("/static/license.xml");
        License license = new License();
        license.setLicense(is);
    }

    public static void main(String[] args) throws Exception {

        List<Test> list = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            Test test = new Test();
            test.setId(Long.valueOf(i));
            test.setIndex("测试字体" + i);
            test.setProblem(String.valueOf(i));
            Abc abc = new Abc(666L, "测试字体");
            test.setList(Arrays.asList(abc));
            list.add(test);
        }
        File file = new File("D:\\25、执勤记录本（100本）.docx");
//        PDFPosition position = Word2PdfBuilder.newInstance(file)
//                .buildParagraphList(list,"list")
//                .defaultTableStyle()
//                .toPDFWithPosition("N");
                InputStream inputStream = Word2PdfBuilder.newInstance(file)
                .buildRowList(list,"tableStyle",3)
                .defaultTableStyle()
                .toDocx();
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\perm.docx");
        IOUtils.copy(inputStream, fileOutputStream);
        fileOutputStream.close();
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Test{

        private Long id;

        private String index;

        private String problem;

        private List<Abc> list;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class Abc{

        private Long id;

        private String index;
    }

}
