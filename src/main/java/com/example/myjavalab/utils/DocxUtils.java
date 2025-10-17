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
            createParagraphBookmark(newParagraph, bookmarkName);
            
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
            createParagraphBookmark(fallbackParagraph, bookmarkName);
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
    private static void createParagraphBookmark(XWPFParagraph paragraph, String bookmarkName) {
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
                    System.out.println("🎯 找到书签 '" + bookmarkName + "'，ID: " + bookmark.getId());
                    // 找到书签，提取书签范围内的内容
                    return extractContentBetweenBookmarks(paragraph, bookmark.getId());
                }
            }
        } catch (Exception e) {
            System.err.println("提取书签内容失败: " + e.getMessage());
        }
        
        // 如果无法提取书签内容，抛出异常，提示书签不存在
        throw new IllegalArgumentException("无法提取书签 '" + bookmarkName + "' 的内容，书签不存在或格式不正确");
    }
    
    /**
     * 提取两个书签标记之间的内容
     * 修复：正确解析XML结构，提取bookmarkStart和bookmarkEnd之间的内容
     * 支持跨段落的书签（bookmarkEnd可能在下一个段落中）
     */
    private static String extractContentBetweenBookmarks(XWPFParagraph paragraph, BigInteger bookmarkId) {
        try {
            CTP ctp = paragraph.getCTP();
            org.w3c.dom.Node paragraphNode = ctp.getDomNode();
            
            // 查找bookmarkStart节点
            org.w3c.dom.Node bookmarkStartNode = findBookmarkStartNode(paragraphNode, bookmarkId);
            if (bookmarkStartNode == null) {
                System.err.println("未找到bookmarkStart节点，ID: " + bookmarkId);
                return "";
            }
            
            // 查找对应的bookmarkEnd节点（可能在当前段落或后续段落中）
            org.w3c.dom.Node bookmarkEndNode = findBookmarkEndNodeInDocument(paragraph, bookmarkId);
            if (bookmarkEndNode == null) {
                System.err.println("未找到bookmarkEnd节点，ID: " + bookmarkId);
                return "";
            }
            
            // 提取两个节点之间的文本内容
            return extractTextBetweenNodes(bookmarkStartNode, bookmarkEndNode);
            
        } catch (Exception e) {
            System.err.println("提取书签内容失败: " + e.getMessage());
            // 如果XML解析失败，回退到段落文本
            String paragraphText = paragraph.getText();
            return paragraphText != null ? paragraphText.trim() : "";
        }
    }
    
    /**
     * 查找指定ID的bookmarkStart节点
     */
    private static org.w3c.dom.Node findBookmarkStartNode(org.w3c.dom.Node paragraphNode, BigInteger bookmarkId) {
        org.w3c.dom.NodeList children = paragraphNode.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            org.w3c.dom.Node child = children.item(i);
            if (child.getLocalName() != null && child.getLocalName().equals("bookmarkStart")) {
                // 检查ID是否匹配
                org.w3c.dom.NamedNodeMap attributes = child.getAttributes();
                if (attributes != null) {
                    org.w3c.dom.Node idAttr = attributes.getNamedItem("w:id");
                    if (idAttr != null) {
                        try {
                            BigInteger nodeId = new BigInteger(idAttr.getNodeValue());
                            if (nodeId.equals(bookmarkId)) {
                                return child;
                            }
                        } catch (NumberFormatException e) {
                            // 忽略格式错误的ID
                        }
                    }
                }
            }
        }
        return null;
    }
    
    /**
     * 查找指定ID的bookmarkEnd节点
     */
    private static org.w3c.dom.Node findBookmarkEndNode(org.w3c.dom.Node paragraphNode, BigInteger bookmarkId) {
        org.w3c.dom.NodeList children = paragraphNode.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            org.w3c.dom.Node child = children.item(i);
            
            // 打印所有子节点信息用于调试
            System.out.println("🔍 检查子节点: " + child.getNodeName() + ", 本地名: " + child.getLocalName() + ", 类型: " + child.getNodeType());
            
            if (child.getLocalName() != null && child.getLocalName().equals("bookmarkEnd")) {
                System.out.println("🎯 找到bookmarkEnd节点！");
                // 检查ID是否匹配
                org.w3c.dom.NamedNodeMap attributes = child.getAttributes();
                if (attributes != null) {
                    // 打印所有属性
                    for (int j = 0; j < attributes.getLength(); j++) {
                        org.w3c.dom.Node attr = attributes.item(j);
                        System.out.println("   属性: " + attr.getNodeName() + " = " + attr.getNodeValue());
                    }
                    
                    org.w3c.dom.Node idAttr = attributes.getNamedItem("w:id");
                    if (idAttr != null) {
                        try {
                            BigInteger nodeId = new BigInteger(idAttr.getNodeValue());
                            System.out.println("🔍 bookmarkEnd节点ID: " + nodeId + ", 查找的ID: " + bookmarkId + ", 匹配: " + nodeId.equals(bookmarkId));
                            if (nodeId.equals(bookmarkId)) {
                                System.out.println("✅ 找到bookmarkEnd节点，ID: " + nodeId + ", 节点名: " + child.getLocalName());
                                return child;
                            }
                        } catch (NumberFormatException e) {
                            // 忽略格式错误的ID
                            System.out.println("⚠️ 无法解析bookmarkEnd ID: " + idAttr.getNodeValue());
                        }
                    } else {
                        System.out.println("⚠️ bookmarkEnd节点没有w:id属性");
                    }
                } else {
                    System.out.println("⚠️ bookmarkEnd节点没有属性");
                }
            }
        }
        return null;
    }
    
    /**
     * 在整个文档中查找指定ID的bookmarkEnd节点
     * 支持跨段落的书签结构，包括段落外的bookmarkEnd节点
     */
    private static org.w3c.dom.Node findBookmarkEndNodeInDocument(XWPFParagraph startParagraph, BigInteger bookmarkId) {
        try {
            System.out.println("🔍 查找bookmarkEnd节点，ID: " + bookmarkId);
            
            // 首先在当前段落中查找
            CTP ctp = startParagraph.getCTP();
            org.w3c.dom.Node paragraphNode = ctp.getDomNode();
            org.w3c.dom.Node bookmarkEndNode = findBookmarkEndNode(paragraphNode, bookmarkId);
            if (bookmarkEndNode != null) {
                System.out.println("✅ 在当前段落找到bookmarkEnd节点");
                return bookmarkEndNode;
            }
            
            // 如果当前段落没找到，在后续段落中查找
            // 获取文档对象
            XWPFDocument document = startParagraph.getDocument();
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            
            // 找到当前段落的索引
            int startIndex = -1;
            for (int i = 0; i < paragraphs.size(); i++) {
                if (paragraphs.get(i) == startParagraph) {
                    startIndex = i;
                    break;
                }
            }
            
            if (startIndex == -1) {
                System.out.println("❌ 找不到当前段落");
                return null; // 找不到当前段落
            }
            
            System.out.println("🔍 在后续段落中查找bookmarkEnd，从段落 " + (startIndex + 1) + " 开始，总段落数: " + paragraphs.size());
            
            // 在后续段落中查找bookmarkEnd
            for (int i = startIndex + 1; i < paragraphs.size(); i++) {
                XWPFParagraph paragraph = paragraphs.get(i);
                CTP paragraphCTP = paragraph.getCTP();
                org.w3c.dom.Node paragraphNode2 = paragraphCTP.getDomNode();
                
                // 打印段落内容用于调试
                String paragraphText = paragraph.getText();
                System.out.println("🔍 检查段落 " + i + ": '" + paragraphText + "'");
                
                bookmarkEndNode = findBookmarkEndNode(paragraphNode2, bookmarkId);
                if (bookmarkEndNode != null) {
                    System.out.println("✅ 在段落 " + i + " 找到bookmarkEnd节点");
                    return bookmarkEndNode;
                }
            }
            
            // 如果没找到，也检查当前段落之前的所有段落
            System.out.println("🔍 检查当前段落之前的所有段落");
            for (int i = 0; i <= startIndex; i++) {
                XWPFParagraph paragraph = paragraphs.get(i);
                CTP paragraphCTP = paragraph.getCTP();
                org.w3c.dom.Node paragraphNode2 = paragraphCTP.getDomNode();
                
                // 打印段落内容用于调试
                String paragraphText = paragraph.getText();
                System.out.println("🔍 检查段落 " + i + ": '" + paragraphText + "'");
                
                bookmarkEndNode = findBookmarkEndNode(paragraphNode2, bookmarkId);
                if (bookmarkEndNode != null) {
                    System.out.println("✅ 在段落 " + i + " 找到bookmarkEnd节点");
                    return bookmarkEndNode;
                }
            }
            
            // 如果段落中都没找到，检查文档主体中的直接子节点
            System.out.println("🔍 检查文档主体中的直接子节点");
            try {
                org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1 documentCT = document.getDocument();
                org.w3c.dom.Node documentNode = documentCT.getDomNode();
                bookmarkEndNode = findBookmarkEndNodeInDocumentBody(documentNode, bookmarkId);
                if (bookmarkEndNode != null) {
                    System.out.println("✅ 在文档主体中找到bookmarkEnd节点");
                    return bookmarkEndNode;
                }
            } catch (Exception e) {
                System.out.println("⚠️ 检查文档主体失败: " + e.getMessage());
            }
            
            System.out.println("❌ 在所有位置都未找到bookmarkEnd节点");
            return null;
        } catch (Exception e) {
            System.err.println("在文档中查找bookmarkEnd节点失败: " + e.getMessage());
            e.printStackTrace();
            return null;
        }
    }
    
    /**
     * 在文档主体中查找bookmarkEnd节点
     */
    private static org.w3c.dom.Node findBookmarkEndNodeInDocumentBody(org.w3c.dom.Node documentNode, BigInteger bookmarkId) {
        org.w3c.dom.NodeList children = documentNode.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            org.w3c.dom.Node child = children.item(i);
            System.out.println("🔍 检查文档主体子节点: " + child.getNodeName() + ", 本地名: " + child.getLocalName());
            
            if (child.getLocalName() != null && child.getLocalName().equals("bookmarkEnd")) {
                System.out.println("🎯 在文档主体中找到bookmarkEnd节点！");
                org.w3c.dom.NamedNodeMap attributes = child.getAttributes();
                if (attributes != null) {
                    org.w3c.dom.Node idAttr = attributes.getNamedItem("w:id");
                    if (idAttr != null) {
                        try {
                            BigInteger nodeId = new BigInteger(idAttr.getNodeValue());
                            System.out.println("🔍 文档主体bookmarkEnd节点ID: " + nodeId + ", 查找的ID: " + bookmarkId + ", 匹配: " + nodeId.equals(bookmarkId));
                            if (nodeId.equals(bookmarkId)) {
                                System.out.println("✅ 在文档主体中找到匹配的bookmarkEnd节点，ID: " + nodeId);
                                return child;
                            }
                        } catch (NumberFormatException e) {
                            System.out.println("⚠️ 无法解析文档主体bookmarkEnd ID: " + idAttr.getNodeValue());
                        }
                    }
                }
            } else if (child.getLocalName() != null && child.getLocalName().equals("body")) {
                // 如果找到body节点，递归搜索其子节点
                System.out.println("🔍 在body节点中递归搜索bookmarkEnd");
                org.w3c.dom.Node result = findBookmarkEndNodeInDocumentBody(child, bookmarkId);
                if (result != null) {
                    return result;
                }
            }
        }
        return null;
    }
    
    /**
     * 提取两个节点之间的文本内容
     * 支持跨段落的书签内容提取
     */
    private static String extractTextBetweenNodes(org.w3c.dom.Node startNode, org.w3c.dom.Node endNode) {
        StringBuilder content = new StringBuilder();
        
        // 如果startNode和endNode在同一个段落中
        if (startNode.getParentNode().equals(endNode.getParentNode())) {
            // 从startNode的下一个兄弟节点开始，到endNode的前一个兄弟节点结束
            org.w3c.dom.Node current = startNode.getNextSibling();
            while (current != null && !current.equals(endNode)) {
                if (current.getNodeType() == org.w3c.dom.Node.TEXT_NODE) {
                    content.append(current.getNodeValue());
                } else if (current.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE) {
                    // 如果是元素节点（如run），提取其中的文本
                    String text = extractTextFromElement(current);
                    if (!text.isEmpty()) {
                        content.append(text);
                    }
                }
                current = current.getNextSibling();
            }
        } else {
            // 跨段落的情况：从startNode开始，到endNode结束
            // 首先提取startNode所在段落中startNode之后的内容
            org.w3c.dom.Node current = startNode.getNextSibling();
            while (current != null) {
                if (current.getNodeType() == org.w3c.dom.Node.TEXT_NODE) {
                    content.append(current.getNodeValue());
                } else if (current.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE) {
                    String text = extractTextFromElement(current);
                    if (!text.isEmpty()) {
                        content.append(text);
                    }
                }
                current = current.getNextSibling();
            }
            
            // 然后提取中间段落的完整内容
            org.w3c.dom.Node startParent = startNode.getParentNode();
            org.w3c.dom.Node endParent = endNode.getParentNode();
            org.w3c.dom.Node currentParent = startParent.getNextSibling();
            
            while (currentParent != null && !currentParent.equals(endParent)) {
                if (currentParent.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE) {
                    String text = extractTextFromElement(currentParent);
                    if (!text.isEmpty()) {
                        content.append(text);
                    }
                }
                currentParent = currentParent.getNextSibling();
            }
            
            // 最后提取endNode所在段落中endNode之前的内容
            current = endParent.getFirstChild();
            while (current != null && !current.equals(endNode)) {
                if (current.getNodeType() == org.w3c.dom.Node.TEXT_NODE) {
                    content.append(current.getNodeValue());
                } else if (current.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE) {
                    String text = extractTextFromElement(current);
                    if (!text.isEmpty()) {
                        content.append(text);
                    }
                }
                current = current.getNextSibling();
            }
        }
        
        return content.toString().trim();
    }
    
    /**
     * 从元素节点中提取文本内容
     */
    private static String extractTextFromElement(org.w3c.dom.Node element) {
        StringBuilder text = new StringBuilder();
        
        if (element.getNodeType() == org.w3c.dom.Node.TEXT_NODE) {
            text.append(element.getNodeValue());
        } else if (element.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE) {
            // 递归提取子节点的文本
            org.w3c.dom.NodeList children = element.getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                text.append(extractTextFromElement(children.item(i)));
            }
        }
        
        return text.toString();
    }
    
    
    /**
     * 替换书签之间的内容，同时保持书签标记不变
     * 使用DOM操作精确替换内容，避免破坏书签结构
     */
    private static void replaceContentBetweenBookmarks(XWPFParagraph paragraph, BigInteger bookmarkId, String newContent) {
        try {
            CTP ctp = paragraph.getCTP();
            org.w3c.dom.Node paragraphNode = ctp.getDomNode();
            
            // 查找bookmarkStart和bookmarkEnd节点
            org.w3c.dom.Node bookmarkStartNode = findBookmarkStartNode(paragraphNode, bookmarkId);
            org.w3c.dom.Node bookmarkEndNode = findBookmarkEndNode(paragraphNode, bookmarkId);
            
            if (bookmarkStartNode == null || bookmarkEndNode == null) {
                System.err.println("无法找到书签标记，ID: " + bookmarkId);
                return;
            }
            
            // 删除bookmarkStart和bookmarkEnd之间的所有内容节点
            removeContentBetweenBookmarks(bookmarkStartNode, bookmarkEndNode);
            
            // 在bookmarkStart之后插入新的内容
            insertContentAfterBookmarkStart(paragraph, bookmarkStartNode, newContent);
            
            System.out.println("✅ 书签内容已替换，ID: " + bookmarkId);
            
        } catch (Exception e) {
            System.err.println("替换书签内容失败: " + e.getMessage());
        }
    }
    
    /**
     * 删除两个书签标记之间的所有内容节点
     */
    private static void removeContentBetweenBookmarks(org.w3c.dom.Node bookmarkStartNode, org.w3c.dom.Node bookmarkEndNode) {
        org.w3c.dom.Node current = bookmarkStartNode.getNextSibling();
        while (current != null && !current.equals(bookmarkEndNode)) {
            org.w3c.dom.Node next = current.getNextSibling();
            // 只删除内容节点，保留书签标记
            if (current.getLocalName() != null && 
                !current.getLocalName().equals("bookmarkStart") && 
                !current.getLocalName().equals("bookmarkEnd")) {
                current.getParentNode().removeChild(current);
            }
            current = next;
        }
    }
    
    /**
     * 在bookmarkStart之后插入新内容
     */
    private static void insertContentAfterBookmarkStart(XWPFParagraph paragraph, org.w3c.dom.Node bookmarkStartNode, String newContent) {
        try {
            // 创建新的run来包含内容
            XWPFRun newRun = paragraph.createRun();
            newRun.setText(newContent);
            
            // 获取新run的DOM节点
            org.w3c.dom.Node newRunNode = newRun.getCTR().getDomNode();
            
            // 将新run插入到bookmarkStart之后
            bookmarkStartNode.getParentNode().insertBefore(newRunNode, bookmarkStartNode.getNextSibling());
            
        } catch (Exception e) {
            System.err.println("插入新内容失败: " + e.getMessage());
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
     * 为书签设置内容并保持编号样式
     * 修复：使用DOM操作保持书签结构，避免破坏bookmarkStart和bookmarkEnd位置
     */
    private static void setBookmarkContentWithNumbering(XWPFDocument document, String bookmarkName, String content, int number) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                try {
                    // 获取书签ID
                    BigInteger bookmarkId = getBookmarkId(paragraph, bookmarkName);
                    if (bookmarkId == null) {
                        System.err.println("无法找到书签ID: " + bookmarkName);
                        break;
                    }
                    
                    // 设置段落为编号列表样式
                    setParagraphNumberingStyle(paragraph, number);
                    
                    // 使用DOM操作替换内容，保持书签结构
                    replaceContentBetweenBookmarks(paragraph, bookmarkId, content);
                    
                    System.out.println("✅ 书签内容已更新，保持书签结构: " + bookmarkName);
                    break;
                    
                } catch (Exception e) {
                    System.err.println("设置书签内容失败: " + e.getMessage());
                    // 如果DOM操作失败，回退到原来的方法
                    fallbackSetBookmarkContent(paragraph, content, number);
                    break;
                }
            }
        }
    }
    
    /**
     * 获取书签的ID
     */
    private static BigInteger getBookmarkId(XWPFParagraph paragraph, String bookmarkName) {
        try {
            CTP ctp = paragraph.getCTP();
            CTBookmark[] bookmarks = ctp.getBookmarkStartArray();
            
            for (CTBookmark bookmark : bookmarks) {
                if (bookmarkName.equals(bookmark.getName())) {
                    return bookmark.getId();
                }
            }
        } catch (Exception e) {
            System.err.println("获取书签ID失败: " + e.getMessage());
        }
        return null;
    }
    
    /**
     * 回退方法：如果DOM操作失败，使用原来的方法
     */
    private static void fallbackSetBookmarkContent(XWPFParagraph paragraph, String content, int number) {
        try {
            // 清除段落中的所有runs
            while (paragraph.getRuns().size() > 0) {
                paragraph.removeRun(0);
            }
            
            // 设置段落为编号列表样式
            setParagraphNumberingStyle(paragraph, number);
            
            // 解析内容并保持样式（不包含序号）
            parseAndSetContentWithStyle(paragraph, content);
            
            System.out.println("⚠️ 使用回退方法设置书签内容");
        } catch (Exception e) {
            System.err.println("回退方法也失败: " + e.getMessage());
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
