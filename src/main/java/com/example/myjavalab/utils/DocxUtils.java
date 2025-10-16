package com.example.myjavalab.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.io.*;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class DocxUtils {

    /**
     * 在指定书签A前面插入新书签B
     * @param inputPath 输入文档路径
     * @param outputPath 输出文档路径
     * @param bookmarkA 目标书签A的名称
     * @param bookmarkB 要插入的书签B的名称
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XmlException
     */
    public static void insertBookmarkBefore(String inputPath, String outputPath, 
                                          String bookmarkA, String bookmarkB) 
                                          throws IOException, InvalidFormatException, XmlException {
        
        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            
            // 查找书签A的位置
            int bookmarkAPosition = findBookmarkPosition(document, bookmarkA);
            if (bookmarkAPosition == -1) {
                throw new IllegalArgumentException("书签 " + bookmarkA + " 未找到");
            }
            
            // 在书签A前面插入书签B
            insertBookmarkAtPosition(document, bookmarkB, bookmarkAPosition);
            
            // 保存文档
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }
        }
    }
    
    /**
     * 将书签A的内容复制到书签B
     * @param inputPath 输入文档路径
     * @param outputPath 输出文档路径
     * @param bookmarkA 源书签A的名称
     * @param bookmarkB 目标书签B的名称
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XmlException
     */
    public static void copyBookmarkContent(String inputPath, String outputPath,
                                        String bookmarkA, String bookmarkB)
                                        throws IOException, InvalidFormatException, XmlException {
        
        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            
            // 获取书签A的内容
            String contentA = getBookmarkContent(document, bookmarkA);
            if (contentA == null) {
                throw new IllegalArgumentException("书签 " + bookmarkA + " 未找到或内容为空");
            }
            
            // 设置书签B的内容
            setBookmarkContent(document, bookmarkB, contentA);
            
            // 保存文档
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }
        }
    }
    
    /**
     * 查找书签在文档中的位置
     */
    private static int findBookmarkPosition(XWPFDocument document, String bookmarkName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph paragraph = paragraphs.get(i);
            if (containsBookmark(paragraph, bookmarkName)) {
                return i;
            }
        }
        return -1;
    }
    
    /**
     * 检查段落是否包含指定的书签
     */
    private static boolean containsBookmark(XWPFParagraph paragraph, String bookmarkName) {
        try {
            CTP ctp = paragraph.getCTP();
            CTBookmark[] bookmarks = ctp.getBookmarkStartArray();
            for (CTBookmark bookmark : bookmarks) {
                if (bookmarkName.equals(bookmark.getName())) {
                    return true;
                }
            }
        } catch (Exception e) {
            // 如果无法访问书签，回退到文本搜索
            String text = paragraph.getText();
            return text != null && text.contains(bookmarkName);
        }
        return false;
    }
    
    /**
     * 在指定位置插入书签
     */
    private static void insertBookmarkAtPosition(XWPFDocument document, String bookmarkName, int position) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        if (position >= 0 && position < paragraphs.size()) {
            XWPFParagraph paragraph = paragraphs.get(position);
            createBookmark(paragraph, bookmarkName);
        }
    }
    
    /**
     * 在段落中创建书签
     */
    private static void createBookmark(XWPFParagraph paragraph, String bookmarkName) {
        try {
            CTP ctp = paragraph.getCTP();
            
            // 创建书签开始标记
            CTBookmark bookmarkStart = ctp.addNewBookmarkStart();
            bookmarkStart.setName(bookmarkName);
            bookmarkStart.setId(BigInteger.valueOf(0));
            
            // 创建书签结束标记
            CTMarkupRange bookmarkEnd = ctp.addNewBookmarkEnd();
            bookmarkEnd.setId(BigInteger.valueOf(0));
            
        } catch (Exception e) {
            System.err.println("创建书签失败: " + e.getMessage());
            // 如果创建书签失败，至少添加文本作为备选
            XWPFRun run = paragraph.createRun();
            run.setText("[" + bookmarkName + "]");
            run.setBold(true);
        }
    }
    
    /**
     * 获取书签的内容
     */
    private static String getBookmarkContent(XWPFDocument document, String bookmarkName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                // 获取书签范围内的内容
                return extractBookmarkContent(paragraph, bookmarkName);
            }
        }
        return null;
    }
    
    /**
     * 从段落中提取书签范围内的内容
     */
    private static String extractBookmarkContent(XWPFParagraph paragraph, String bookmarkName) {
        try {
            CTP ctp = paragraph.getCTP();
            CTBookmark[] bookmarks = ctp.getBookmarkStartArray();
            
            for (CTBookmark bookmark : bookmarks) {
                if (bookmarkName.equals(bookmark.getName())) {
                    // 找到书签，提取书签范围内的内容
                    return extractContentBetweenBookmarks(paragraph, bookmark.getId());
                }
            }
        } catch (Exception e) {
            System.err.println("提取书签内容失败: " + e.getMessage());
        }
        
        // 如果无法提取书签内容，返回整个段落文本作为备选
        String paragraphText = paragraph.getText();
        return paragraphText != null ? paragraphText.trim() : "";
    }
    
    /**
     * 提取两个书签标记之间的内容
     */
    private static String extractContentBetweenBookmarks(XWPFParagraph paragraph, BigInteger bookmarkId) {
        // 简化实现：由于书签内容提取比较复杂，暂时返回整个段落文本
        // 在实际应用中，这可能需要更复杂的XML解析逻辑
        String paragraphText = paragraph.getText();
        if (paragraphText != null) {
            // 尝试从段落文本中提取书签内容
            // 这是一个简化的实现，实际应该解析XML结构
            return paragraphText.trim();
        }
        return "";
    }
    
    /**
     * 设置书签的内容
     */
    private static void setBookmarkContent(XWPFDocument document, String bookmarkName, String content) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                // 简化实现：清除段落中的所有runs，然后添加新内容
                while (paragraph.getRuns().size() > 0) {
                    paragraph.removeRun(0);
                }
                
                // 创建新的run并设置内容
                XWPFRun run = paragraph.createRun();
                run.setText(content);
                break;
            }
        }
    }
    
    /**
     * 获取文档中指定书签的内容（公共方法，用于测试验证）
     * @param documentPath 文档路径
     * @param bookmarkName 书签名称
     * @return 书签内容，如果未找到返回null
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XmlException
     */
    public static String getBookmarkContentFromFile(String documentPath, String bookmarkName) 
                                                   throws IOException, InvalidFormatException, XmlException {
        try (FileInputStream fis = new FileInputStream(documentPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            return getBookmarkContent(document, bookmarkName);
        }
    }

    /**
     * 对指定书签进行多次内容复制
     * @param sourceFile 需要操作的源文件
     * @param targetFile 原文件操作的结果的存储文件
     * @param sourceLabel 需要执行内容复制操作的书签
     * @param copyTimes 书签内容复制次数
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XmlException
     */
    public static void copyBookmarkContentMultipleTimes(String sourceFile, String targetFile, 
                                                      String sourceLabel, int copyTimes) 
                                                      throws IOException, InvalidFormatException, XmlException {
        
        // 构建完整的源文件路径
        String sourcePath = "src/main/resources/doc/" + sourceFile;
        
        try (FileInputStream fis = new FileInputStream(sourcePath);
             XWPFDocument document = new XWPFDocument(fis)) {
            
            // 获取源书签的内容
            String sourceContent = getBookmarkContent(document, sourceLabel);
            if (sourceContent == null) {
                throw new IllegalArgumentException("书签 " + sourceLabel + " 未找到或内容为空");
            }
            
            // 找到源书签的位置
            int sourcePosition = findBookmarkPosition(document, sourceLabel);
            if (sourcePosition == -1) {
                throw new IllegalArgumentException("书签 " + sourceLabel + " 未找到");
            }
            
            // 在源书签之前插入多个新书签并复制内容
            for (int i = 1; i <= copyTimes; i++) {
                String targetLabel = sourceLabel + i;
                
                // 在源书签之前插入新书签
                insertBookmarkAtPosition(document, targetLabel, sourcePosition);
                
                // 将源书签的内容复制给新书签
                setBookmarkContent(document, targetLabel, sourceContent);
                
                System.out.println("✅ 已创建书签 " + targetLabel + " 并复制内容");
            }
            
            // 保存文档到doc目录
            String outputPath = "src/main/resources/doc/" + targetFile;
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }
            
            System.out.println("✅ 文档已保存到: " + outputPath);
            System.out.println("📊 总共创建了 " + copyTimes + " 个新书签");
        }
    }
}
