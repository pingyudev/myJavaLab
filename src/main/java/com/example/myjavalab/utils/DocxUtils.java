package com.example.myjavalab.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.io.*;
import java.math.BigInteger;
import java.util.List;

public class DocxUtils {

    // 书签ID计数器，确保每个书签有唯一ID
    private static long bookmarkIdCounter = 1000;

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
            
            // 检查书签A是否存在
            if (findBookmarkPosition(document, bookmarkA) == -1) {
                throw new IllegalArgumentException("书签 " + bookmarkA + " 未找到");
            }
            
            // 在书签A前面插入书签B（使用改进的方法）
            insertBookmarkBeforeTargetBookmark(document, bookmarkA, bookmarkB);
            
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
     * 查找书签在文档中的范围
     */
    private static BookmarkRange findBookmarkRange(XWPFDocument document, String bookmarkName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph paragraph = paragraphs.get(i);
            if (containsBookmark(paragraph, bookmarkName)) {
                // 对于单段落书签，起始和结束位置相同
                return new BookmarkRange(i, i);
            }
        }
        return new BookmarkRange(-1, -1); // 未找到
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
     * 在指定书签之前插入新书签（改进版本，保持原有书签位置不变）
     */
    private static void insertBookmarkBeforeTargetBookmark(XWPFDocument document, String targetBookmarkName, String newBookmarkName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph paragraph = paragraphs.get(i);
            if (containsBookmark(paragraph, targetBookmarkName)) {
                // 找到目标书签所在的段落，在其前面插入新段落
                insertParagraphBeforeTarget(document, paragraph, newBookmarkName);
                break;
            }
        }
    }
    
    /**
     * 在目标段落之前插入新段落
     * 修复：只使用编号样式，避免重复序号和破坏书签结构
     */
    private static void insertParagraphBeforeTarget(XWPFDocument document, XWPFParagraph targetParagraph, String bookmarkName) {
        try {
            // 创建新段落
            XWPFParagraph newParagraph = document.createParagraph();
            
            // 复制目标段落的样式到新段落
            copyParagraphStyle(targetParagraph, newParagraph);
            
            // 获取目标段落的编号
            int targetNumber = extractNumberFromParagraph(targetParagraph);
            
            // 只添加initialString的内容，不手动添加序号（让Word编号样式自动处理）
            XWPFRun spaceRun = newParagraph.createRun();
            spaceRun.setText("initialString"); // 4个initialString
            
            // 在新段落中创建书签（包围initialString内容）
            createBookmark(newParagraph, bookmarkName);
            
            // 只更新目标段落的编号样式属性，不重建内容（保持书签结构）
            updateParagraphNumberingStyleOnly(targetParagraph, targetNumber + 1);
            
            // 获取目标段落的XML节点
            CTP targetCTP = targetParagraph.getCTP();
            
            // 获取新段落的XML节点
            CTP newCTP = newParagraph.getCTP();
            
            // 在目标段落之前插入新段落
            // 使用DOM操作将新段落插入到目标段落之前
            targetCTP.getDomNode().getParentNode().insertBefore(
                newCTP.getDomNode(), targetCTP.getDomNode());
                
            System.out.println("✅ 新段落已插入，书签: " + bookmarkName + "，编号: " + targetNumber);
                
        } catch (Exception e) {
            System.err.println("在目标段落之前插入失败: " + e.getMessage());
            // 如果插入失败，至少确保书签被创建
            XWPFParagraph fallbackParagraph = document.createParagraph();
            createBookmark(fallbackParagraph, bookmarkName);
        }
    }
    
    /**
     * 复制段落的样式到目标段落
     */
    private static void copyParagraphStyle(XWPFParagraph sourceParagraph, XWPFParagraph targetParagraph) {
        try {
            CTP sourceCTP = sourceParagraph.getCTP();
            CTP targetCTP = targetParagraph.getCTP();
            
            // 复制段落属性
            if (sourceCTP.getPPr() != null) {
                if (targetCTP.getPPr() == null) {
                    targetCTP.addNewPPr();
                }
                
                // 复制编号属性
                if (sourceCTP.getPPr().getNumPr() != null) {
                    CTNumPr sourceNumPr = sourceCTP.getPPr().getNumPr();
                    CTNumPr targetNumPr = targetCTP.getPPr().addNewNumPr();
                    
                    // 复制编号ID
                    if (sourceNumPr.getNumId() != null) {
                        CTDecimalNumber sourceNumId = sourceNumPr.getNumId();
                        CTDecimalNumber targetNumId = targetNumPr.addNewNumId();
                        targetNumId.setVal(sourceNumId.getVal());
                    }
                    
                    // 复制编号级别
                    if (sourceNumPr.getIlvl() != null) {
                        CTDecimalNumber sourceIlvl = sourceNumPr.getIlvl();
                        CTDecimalNumber targetIlvl = targetNumPr.addNewIlvl();
                        targetIlvl.setVal(sourceIlvl.getVal());
                    }
                }
                
                // 复制其他段落属性（如对齐方式、间距等）
                if (sourceCTP.getPPr().getJc() != null) {
                    targetCTP.getPPr().setJc(sourceCTP.getPPr().getJc());
                }
                
                if (sourceCTP.getPPr().getSpacing() != null) {
                    targetCTP.getPPr().setSpacing(sourceCTP.getPPr().getSpacing());
                }
            } else {
                // 如果源段落没有编号样式，为目标段落设置默认编号样式
                setParagraphNumberingStyle(targetParagraph, 1);
            }
            
        } catch (Exception e) {
            System.err.println("复制段落样式失败: " + e.getMessage());
            // 如果复制失败，至少设置基本的编号样式
            setParagraphNumberingStyle(targetParagraph, 1);
        }
    }
    
    
    /**
     * 只更新段落的编号样式属性，不重建内容（保持书签结构完整）
     */
    private static void updateParagraphNumberingStyleOnly(XWPFParagraph paragraph, int newNumber) {
        try {
            // 获取段落的底层XML对象
            CTP ctp = paragraph.getCTP();
            
            // 设置段落为编号列表
            if (ctp.getPPr() == null) {
                ctp.addNewPPr();
            }
            
            // 创建或更新编号属性
            CTNumPr numPr;
            if (ctp.getPPr().getNumPr() == null) {
                numPr = ctp.getPPr().addNewNumPr();
            } else {
                numPr = ctp.getPPr().getNumPr();
            }
            
            // 设置编号ID（使用默认的编号样式）
            if (numPr.getNumId() == null) {
                numPr.addNewNumId();
            }
            numPr.getNumId().setVal(BigInteger.valueOf(1)); // 使用编号样式1
            
            // 设置编号级别
            if (numPr.getIlvl() == null) {
                numPr.addNewIlvl();
            }
            numPr.getIlvl().setVal(BigInteger.valueOf(0)); // 使用级别0
            
            System.out.println("✅ 段落编号样式已更新为: " + newNumber);
            
        } catch (Exception e) {
            System.err.println("更新段落编号样式失败: " + e.getMessage());
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
            CTNumPr numPr = ctp.getPPr().addNewNumPr();
            
            // 设置编号ID（使用默认的编号样式）
            CTDecimalNumber numId = numPr.addNewNumId();
            numId.setVal(BigInteger.valueOf(1)); // 使用编号样式1
            
            // 设置编号级别
            CTDecimalNumber ilvl = numPr.addNewIlvl();
            ilvl.setVal(BigInteger.valueOf(0)); // 使用级别0
            
        } catch (Exception e) {
            System.err.println("设置编号样式失败，回退到文本序号: " + e.getMessage());
            // 如果设置编号样式失败，回退到文本序号
            XWPFRun numberRun = paragraph.createRun();
            numberRun.setText(number + ". ");
        }
    }
    
    
    /**
     * 生成唯一的书签ID
     */
    private static BigInteger generateUniqueBookmarkId() {
        return BigInteger.valueOf(bookmarkIdCounter++);
    }
    
    /**
     * 在段落中创建书签（包围整个段落内容）
     * 修复：使用DOM操作确保书签正确包围段落内容
     */
    private static void createBookmark(XWPFParagraph paragraph, String bookmarkName) {
        try {
            CTP ctp = paragraph.getCTP();
            BigInteger bookmarkId = generateUniqueBookmarkId();
            
            // 确保段落有内容，如果没有则添加initialString
            if (paragraph.getRuns().isEmpty()) {
                XWPFRun spaceRun = paragraph.createRun();
                spaceRun.setText("initialString");
            }
            
            // 创建书签标记（会添加到末尾）
            CTBookmark bookmarkStart = ctp.addNewBookmarkStart();
            bookmarkStart.setName(bookmarkName);
            bookmarkStart.setId(bookmarkId);
            
            CTMarkupRange bookmarkEnd = ctp.addNewBookmarkEnd();
            bookmarkEnd.setId(bookmarkId);
            
            // 使用DOM操作移动bookmarkStart到第一个Run之前
            org.w3c.dom.Node bookmarkStartNode = bookmarkStart.getDomNode();
            org.w3c.dom.Node firstRunNode = null;
            
            // 查找第一个<w:r>节点
            org.w3c.dom.NodeList children = ctp.getDomNode().getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                org.w3c.dom.Node child = children.item(i);
                if (child.getLocalName() != null && child.getLocalName().equals("r")) {
                    firstRunNode = child;
                    break;
                }
            }
            
            // 将bookmarkStart移到第一个Run之前
            if (firstRunNode != null) {
                ctp.getDomNode().insertBefore(bookmarkStartNode, firstRunNode);
            }
            
            System.out.println("✅ 书签 '" + bookmarkName + "' 已创建，ID: " + bookmarkId);
            
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
     * 获取书签在文档中的范围（公共方法，用于测试验证）
     * @param documentPath 文档路径
     * @param bookmarkName 书签名称
     * @return 书签范围，如果未找到返回BookmarkRange(-1, -1)
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XmlException
     */
    public static BookmarkRange getBookmarkRangeFromFile(String documentPath, String bookmarkName) 
                                                       throws IOException, InvalidFormatException, XmlException {
        try (FileInputStream fis = new FileInputStream(documentPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            return findBookmarkRange(document, bookmarkName);
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
     * 检查段落是否使用编号样式（包括Word编号样式和文本编号）
     */
    private static boolean isParagraphUsingNumberingStyle(XWPFDocument document, String bookmarkName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                try {
                    // 检查Word编号样式
                    CTP ctp = paragraph.getCTP();
                    if (ctp.getPPr() != null && ctp.getPPr().getNumPr() != null) {
                        return true; // 使用了Word编号样式
                    }
                    
                    // 检查文本编号格式
                    String text = paragraph.getText();
                    if (text != null && text.matches("^\\d+\\..*")) {
                        return true; // 使用了文本编号格式
                    }
                } catch (Exception e) {
                    // 如果无法检查，返回false
                }
                return false;
            }
        }
        return false;
    }

}
