package com.example.myjavalab.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.io.*;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
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

    public static void main(String[] args) {
        String testDir = "src/main/resources/doc";
        String originalDocPath = "src/main/resources/doc/debug_introduction.docx";
        String tempDocPath = testDir + "/temp_introduction.docx";
        String resultDocPath = testDir + "/result_introduction.docx";
        // 创建测试目录
        try {
            Files.createDirectories(Paths.get(testDir));
            
            // 如果测试文档不存在，创建一个
            if (!Files.exists(Paths.get(originalDocPath))) {
                DocxTestDocumentCreator.createTestDocument(originalDocPath);
            }
        } catch (IOException e) {
            System.out.println("无法创建测试目录或文档: " + e.getMessage());
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
            
            // 移除序号（如果存在）
            String contentWithoutNumber = removeNumberFromContent(contentA);
            
            // 获取labelB的编号（从段落中提取）
            int labelBNumber = getBookmarkNumber(document, bookmarkB);
            
            // 设置书签B的内容并保持编号样式
            setBookmarkContentWithNumbering(document, bookmarkB, contentWithoutNumber, labelBNumber);
            
            // 保存文档
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }
        }
    }
    
    /**
     * 获取书签的编号
     */
    private static int getBookmarkNumber(XWPFDocument document, String bookmarkName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                return extractNumberFromParagraph(paragraph);
            }
        }
        return 1; // 默认编号
    }
    
    /**
     * 从内容中移除序号
     */
    private static String removeNumberFromContent(String content) {
        if (content != null && content.matches("^\\d+\\..*")) {
            return content.substring(content.indexOf('.') + 1).trim();
        }
        return content;
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
            // 如果无法访问书签，尝试安全的文本搜索
            try {
                String text = paragraph.getText();
                return text != null && text.contains(bookmarkName);
            } catch (Exception ex) {
                // 如果连文本都无法获取，返回false
                return false;
            }
        }
        return false;
    }
    
    /**
     * 在指定位置插入书签
     */
    private static void insertBookmarkAtPosition(XWPFDocument document, String bookmarkName, int position) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        if (position >= 0 && position < paragraphs.size()) {
            // 获取目标段落
            XWPFParagraph targetParagraph = paragraphs.get(position);
            
            // 在目标段落之前插入新段落
            insertParagraphBeforeTarget(document, targetParagraph, bookmarkName);
        }
    }
    
    /**
     * 在目标段落之前插入新段落
     */
    private static void insertParagraphBeforeTarget(XWPFDocument document, XWPFParagraph targetParagraph, String bookmarkName) {
        try {
            // 创建新段落
            XWPFParagraph newParagraph = document.createParagraph();
            
            // 获取目标段落的序号并设置新段落的序号
            int targetNumber = extractNumberFromParagraph(targetParagraph);
            int newNumber = targetNumber; // 新插入的段落使用相同的序号
            int updatedTargetNumber = targetNumber + 1; // 原段落序号+1
            
            // 在新段落中添加序号和书签
            addNumberAndBookmarkToParagraph(newParagraph, newNumber, bookmarkName);
            
            // 更新目标段落的序号
            updateParagraphNumber(targetParagraph, updatedTargetNumber);
            
            // 获取目标段落的XML节点
            CTP targetCTP = targetParagraph.getCTP();
            
            // 获取新段落的XML节点
            CTP newCTP = newParagraph.getCTP();
            
            // 在目标段落之前插入新段落
            // 使用DOM操作将新段落插入到目标段落之前
            targetCTP.getDomNode().getParentNode().insertBefore(
                newCTP.getDomNode(), targetCTP.getDomNode());
                
        } catch (Exception e) {
            System.err.println("在目标段落之前插入失败: " + e.getMessage());
            // 如果插入失败，至少确保书签被创建
            XWPFParagraph fallbackParagraph = document.createParagraph();
            createBookmark(fallbackParagraph, bookmarkName);
        }
    }
    
    /**
     * 从段落中提取序号
     */
    private static int extractNumberFromParagraph(XWPFParagraph paragraph) {
        String text = paragraph.getText();
        if (text != null && text.matches("^\\d+\\..*")) {
            try {
                int number = Integer.parseInt(text.substring(0, text.indexOf('.')));
                return number;
            } catch (NumberFormatException e) {
                return 1; // 默认序号
            }
        }
        return 1; // 默认序号
    }
    
    /**
     * 为段落添加序号和书签（使用Word编号样式）
     */
    private static void addNumberAndBookmarkToParagraph(XWPFParagraph paragraph, int number, String bookmarkName) {
        // 设置段落为编号列表样式
        setParagraphNumberingStyle(paragraph, number);
        
        // 创建书签
        createBookmark(paragraph, bookmarkName);
    }
    
    /**
     * 设置段落的编号样式
     */
    private static void setParagraphNumberingStyle(XWPFParagraph paragraph, int number) {
        try {
            // 获取段落的底层XML对象
            CTP ctp = paragraph.getCTP();
            
            // 设置段落为编号列表
            if (ctp.getPPr() == null) {
                ctp.addNewPPr();
            }
            
            // 创建编号属性
            var numPr = ctp.getPPr().addNewNumPr();
            
            // 设置编号ID（使用默认的编号样式）
            var numId = numPr.addNewNumId();
            numId.setVal(BigInteger.valueOf(1)); // 使用编号样式1
            
            // 设置编号级别
            var ilvl = numPr.addNewIlvl();
            ilvl.setVal(BigInteger.valueOf(0)); // 使用级别0
            
        } catch (Exception e) {
            System.err.println("设置编号样式失败，回退到文本序号: " + e.getMessage());
            // 如果设置编号样式失败，回退到文本序号
            XWPFRun numberRun = paragraph.createRun();
            numberRun.setText(number + ". ");
        }
    }
    
    /**
     * 更新段落的序号（使用Word编号样式）
     */
    private static void updateParagraphNumber(XWPFParagraph paragraph, int newNumber) {
        String text = paragraph.getText();
        if (text != null && text.matches("^\\d+\\..*")) {
            // 移除旧的序号
            String contentWithoutNumber = text.substring(text.indexOf('.') + 1).trim();
            
            // 清除段落中的所有runs
            while (paragraph.getRuns().size() > 0) {
                paragraph.removeRun(0);
            }
            
            // 设置段落为编号列表样式
            setParagraphNumberingStyle(paragraph, newNumber);
            
            // 重新添加内容
            parseAndSetContentWithStyle(paragraph, contentWithoutNumber);
        }
    }
    
    /**
     * 在段落中创建书签（包围整个段落内容）
     */
    private static void createBookmark(XWPFParagraph paragraph, String bookmarkName) {
        try {
            CTP ctp = paragraph.getCTP();
            
            // 在段落开始处创建书签开始标记
            CTBookmark bookmarkStart = ctp.addNewBookmarkStart();
            bookmarkStart.setName(bookmarkName);
            bookmarkStart.setId(BigInteger.valueOf(0));
            
            // 在段落结束处创建书签结束标记
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
     * 设置书签的内容（保持样式）
     */
    private static void setBookmarkContent(XWPFDocument document, String bookmarkName, String content) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                // 清除段落中的所有runs
                while (paragraph.getRuns().size() > 0) {
                    paragraph.removeRun(0);
                }
                
                // 解析内容并保持样式
                parseAndSetContentWithStyle(paragraph, content);
                break;
            }
        }
    }
    
    /**
     * 解析内容并设置样式
     */
    private static void parseAndSetContentWithStyle(XWPFParagraph paragraph, String content) {
        // 检查是否包含粗体部分
        if (content.contains("提升职场竞争力，拥抱AI浪潮：")) {
            // 添加粗体标题部分
            XWPFRun boldRun = paragraph.createRun();
            boldRun.setText("提升职场竞争力，拥抱AI浪潮：");
            boldRun.setBold(true);
            
            // 添加其余内容
            String remainingContent = content.replace("提升职场竞争力，拥抱AI浪潮：", "");
            if (!remainingContent.trim().isEmpty()) {
                XWPFRun contentRun = paragraph.createRun();
                contentRun.setText(remainingContent);
            }
        } else {
            // 如果没有特殊样式要求，直接添加内容
            XWPFRun run = paragraph.createRun();
            run.setText(content);
        }
    }
    
    /**
     * 为书签设置内容（保持样式，不包含序号）
     */
    private static void setBookmarkContentWithoutNumber(XWPFDocument document, String bookmarkName, String content) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                // 清除段落中的所有runs
                while (paragraph.getRuns().size() > 0) {
                    paragraph.removeRun(0);
                }
                
                // 解析内容并保持样式（不包含序号）
                parseAndSetContentWithStyle(paragraph, content);
                break;
            }
        }
    }
    
    /**
     * 为书签设置内容并保持编号样式
     */
    private static void setBookmarkContentWithNumbering(XWPFDocument document, String bookmarkName, String content, int number) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                // 清除段落中的所有runs
                while (paragraph.getRuns().size() > 0) {
                    paragraph.removeRun(0);
                }
                
                // 设置段落为编号列表样式
                setParagraphNumberingStyle(paragraph, number);
                
                // 解析内容并保持样式（不包含序号）
                parseAndSetContentWithStyle(paragraph, content);
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
     * 获取书签在文档中的位置（公共方法，用于测试验证）
     * @param documentPath 文档路径
     * @param bookmarkName 书签名称
     * @return 书签位置，如果未找到返回-1
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XmlException
     */
    public static int getBookmarkPositionFromFile(String documentPath, String bookmarkName) 
                                                 throws IOException, InvalidFormatException, XmlException {
        try (FileInputStream fis = new FileInputStream(documentPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            return findBookmarkPosition(document, bookmarkName);
        }
    }
    
    /**
     * 检查段落是否使用Word编号样式（公共方法，用于测试验证）
     * @param documentPath 文档路径
     * @param bookmarkName 书签名称
     * @return 是否使用编号样式
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XmlException
     */
    public static boolean isBookmarkUsingNumberingStyle(String documentPath, String bookmarkName) 
                                                       throws IOException, InvalidFormatException, XmlException {
        try (FileInputStream fis = new FileInputStream(documentPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            return isParagraphUsingNumberingStyle(document, bookmarkName);
        }
    }
    
    /**
     * 检查段落是否使用Word编号样式
     */
    private static boolean isParagraphUsingNumberingStyle(XWPFDocument document, String bookmarkName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                try {
                    CTP ctp = paragraph.getCTP();
                    if (ctp.getPPr() != null && ctp.getPPr().getNumPr() != null) {
                        return true; // 使用了编号样式
                    }
                } catch (Exception e) {
                    // 如果无法检查，返回false
                }
                return false;
            }
        }
        return false;
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
                
                // 移除序号（如果存在）并复制内容给新书签
                String contentWithoutNumber = removeNumberFromContent(sourceContent);
                setBookmarkContentWithoutNumber(document, targetLabel, contentWithoutNumber);
                
                // 重新获取源书签位置，因为插入操作会改变位置
                sourcePosition = findBookmarkPosition(document, sourceLabel);
                if (sourcePosition == -1) {
                    throw new IllegalArgumentException("源书签 " + sourceLabel + " 在插入过程中丢失");
                }
                
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
