package com.example.myjavalab.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.*;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class DocxUtils {

    // 书签ID计数器，确保每个书签有唯一ID
    private static long bookmarkIdCounter = 1000;
    
    /**
     * 段落内容类，用于保存段落的结构信息
     */
    public static class ParagraphContent {
        private final int paragraphIndex;
        private final List<Node> runNodes;
        private final CTP paragraphProperties;
        
        public ParagraphContent(int paragraphIndex, List<Node> runNodes, CTP paragraphProperties) {
            this.paragraphIndex = paragraphIndex;
            this.runNodes = runNodes;
            this.paragraphProperties = paragraphProperties;
        }
        
        public int getParagraphIndex() { return paragraphIndex; }
        public List<Node> getRunNodes() { return runNodes; }
        public CTP getParagraphProperties() { return paragraphProperties; }
    }

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
            
            // 获取书签A的段落内容（支持多段落书签）
            List<ParagraphContent> paragraphContentsA = getBookmarkParagraphContent(document, bookmarkA);
            if (paragraphContentsA.isEmpty()) {
                throw new IllegalArgumentException("书签 " + bookmarkA + " 未找到或内容为空");
            }
            
            // 设置书签B的内容，保持段落结构
            setBookmarkContentFromParagraphContent(document, bookmarkB, paragraphContentsA);
            
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
     * 查找书签在文档中的范围
     * 支持单段落和多段落书签
     */
    private static BookmarkRange findBookmarkRange(XWPFDocument document, String bookmarkName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph paragraph = paragraphs.get(i);
            if (containsBookmark(paragraph, bookmarkName)) {
                // 找到书签起始段落，现在需要找到结束段落
                BigInteger bookmarkId = getBookmarkId(paragraph, bookmarkName);
                if (bookmarkId == null) {
                    return new BookmarkRange(-1, -1); // 无法获取书签ID
                }
                
                // 查找bookmarkEnd节点来确定结束段落
                Node bookmarkEndNode = findBookmarkEndNodeInDocument(paragraph, bookmarkId);
                if (bookmarkEndNode == null) {
                    // 如果找不到bookmarkEnd，假设是单段落书签
                    return new BookmarkRange(i, i);
                }
                
                // 确定bookmarkEnd所在的段落索引
                int endParagraphIndex = findParagraphIndexContainingNode(document, bookmarkEndNode);
                System.out.println("🔍 书签 " + bookmarkName + " 起始段落: " + i + ", 结束段落: " + endParagraphIndex);
                if (endParagraphIndex == -1) {
                    // 如果无法确定结束段落，假设是单段落书签
                    return new BookmarkRange(i, i);
                }
                
                // 确保start <= end
                if (i <= endParagraphIndex) {
                    return new BookmarkRange(i, endParagraphIndex);
                } else {
                    // 如果end < start，交换它们
                    System.out.println("⚠️ 书签范围异常，交换start和end: " + i + " -> " + endParagraphIndex);
                    return new BookmarkRange(endParagraphIndex, i);
                }
            }
        }
        return new BookmarkRange(-1, -1); // 未找到
    }
    
    /**
     * 查找包含指定书签的段落
     */
    private static XWPFParagraph findParagraphWithBookmark(XWPFDocument document, String bookmarkName) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            if (containsBookmark(paragraph, bookmarkName)) {
                return paragraph;
            }
        }
        return null;
    }
    
    /**
     * 查找包含指定DOM节点的段落索引
     */
    private static int findParagraphIndexContainingNode(XWPFDocument document, Node targetNode) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph paragraph = paragraphs.get(i);
            CTP ctp = paragraph.getCTP();
            Node paragraphNode = ctp.getDomNode();
            
            // 检查目标节点是否在当前段落中
            if (isNodeContainedIn(paragraphNode, targetNode)) {
                return i;
            }
        }
        return -1; // 未找到
    }
    
    /**
     * 检查目标节点是否包含在指定段落节点中
     */
    private static boolean isNodeContainedIn(Node paragraphNode, Node targetNode) {
        // 如果目标节点就是段落节点本身
        if (paragraphNode.equals(targetNode)) {
            return true;
        }
        
        // 递归检查子节点
        NodeList children = paragraphNode.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if (child.equals(targetNode) || isNodeContainedIn(child, targetNode)) {
                return true;
            }
        }
        
        return false;
    }
    
    /**
     * 获取书签的run节点（包含格式信息）
     */
    private static List<Node> getBookmarkRunNodes(XWPFDocument document, String bookmarkName) {
        XWPFParagraph paragraph = findParagraphWithBookmark(document, bookmarkName);
        if (paragraph == null) {
            return new ArrayList<>();
        }
        
        BigInteger bookmarkId = getBookmarkId(paragraph, bookmarkName);
        if (bookmarkId == null) {
            return new ArrayList<>();
        }
        
        return extractRunNodesBetweenBookmarks(paragraph, bookmarkId);
    }
    
    /**
     * 获取书签的段落内容（支持多段落书签）
     */
    private static List<ParagraphContent> getBookmarkParagraphContent(XWPFDocument document, String bookmarkName) {
        XWPFParagraph paragraph = findParagraphWithBookmark(document, bookmarkName);
        if (paragraph == null) {
            return new ArrayList<>();
        }
        
        BigInteger bookmarkId = getBookmarkId(paragraph, bookmarkName);
        if (bookmarkId == null) {
            return new ArrayList<>();
        }
        
        // 查找bookmarkStart和bookmarkEnd节点
        CTP ctp = paragraph.getCTP();
        Node paragraphNode = ctp.getDomNode();
        Node bookmarkStartNode = findBookmarkStartNode(paragraphNode, bookmarkId);
        Node bookmarkEndNode = findBookmarkEndNodeInDocument(paragraph, bookmarkId);
        
        if (bookmarkStartNode == null || bookmarkEndNode == null) {
            return new ArrayList<>();
        }
        
        return extractParagraphContentBetweenBookmarks(document, bookmarkStartNode, bookmarkEndNode);
    }
    
    /**
     * 提取书签之间的段落内容（支持多段落书签）
     * 返回按段落组织的结构信息，保持段落边界
     */
    private static List<ParagraphContent> extractParagraphContentBetweenBookmarks(XWPFDocument document, 
                                                                                 Node bookmarkStartNode, 
                                                                                 Node bookmarkEndNode) {
        List<ParagraphContent> paragraphContents = new ArrayList<>();
        
        try {
            // 如果bookmarkStart和bookmarkEnd在同一个段落中
            if (bookmarkStartNode.getParentNode().equals(bookmarkEndNode.getParentNode())) {
                // 单段落情况：提取run节点
                List<Node> runNodes = new ArrayList<>();
                Node current = bookmarkStartNode.getNextSibling();
                while (current != null && !current.equals(bookmarkEndNode)) {
                    if (current.getNodeType() == Node.ELEMENT_NODE && 
                        current.getLocalName() != null && 
                        current.getLocalName().equals("r")) {
                        runNodes.add(current);
                    }
                    current = current.getNextSibling();
                }
                
                // 获取段落属性
                XWPFParagraph paragraph = findParagraphContainingNode(document, bookmarkStartNode);
                CTP paragraphProps = paragraph != null ? paragraph.getCTP() : null;
                paragraphContents.add(new ParagraphContent(0, runNodes, paragraphProps));
                
            } else {
                // 多段落情况：按段落组织内容
                Node startParent = bookmarkStartNode.getParentNode();
                Node endParent = bookmarkEndNode.getParentNode();
                
                // 获取文档的段落列表
                List<XWPFParagraph> paragraphs = document.getParagraphs();
                int startParagraphIndex = findParagraphIndexContainingNode(document, startParent);
                int endParagraphIndex = findParagraphIndexContainingNode(document, endParent);
                
                if (startParagraphIndex != -1 && endParagraphIndex != -1) {
                    // 处理每个段落
                    for (int i = startParagraphIndex; i <= endParagraphIndex; i++) {
                        XWPFParagraph paragraph = paragraphs.get(i);
                        CTP ctp = paragraph.getCTP();
                        Node paragraphNode = ctp.getDomNode();
                        List<Node> runNodes = new ArrayList<>();
                        
                        if (i == startParagraphIndex) {
                            // 起始段落：提取bookmarkStart之后的所有run节点
                            Node current = bookmarkStartNode.getNextSibling();
                            while (current != null) {
                                if (current.getNodeType() == Node.ELEMENT_NODE && 
                                    current.getLocalName() != null && 
                                    current.getLocalName().equals("r")) {
                                    runNodes.add(current);
                                }
                                current = current.getNextSibling();
                            }
                        } else if (i == endParagraphIndex) {
                            // 结束段落：提取bookmarkEnd之前的所有run节点
                            Node current = paragraphNode.getFirstChild();
                            while (current != null && !current.equals(bookmarkEndNode)) {
                                if (current.getNodeType() == Node.ELEMENT_NODE && 
                                    current.getLocalName() != null && 
                                    current.getLocalName().equals("r")) {
                                    runNodes.add(current);
                                }
                                current = current.getNextSibling();
                            }
                        } else {
                            // 中间段落：提取整个段落的所有run节点
                            NodeList children = paragraphNode.getChildNodes();
                            for (int j = 0; j < children.getLength(); j++) {
                                Node child = children.item(j);
                                if (child.getNodeType() == Node.ELEMENT_NODE && 
                                    child.getLocalName() != null && 
                                    child.getLocalName().equals("r")) {
                                    runNodes.add(child);
                                }
                            }
                        }
                        
                        // 创建段落内容对象
                        paragraphContents.add(new ParagraphContent(i - startParagraphIndex, runNodes, ctp));
                    }
                }
            }
            
            System.out.println("✅ 提取到 " + paragraphContents.size() + " 个段落的内容，保持段落结构");
            
        } catch (Exception e) {
            System.err.println("提取段落内容失败: " + e.getMessage());
        }
        
        return paragraphContents;
    }
    
    /**
     * 提取书签之间的段落节点（支持多段落书签）
     * 返回包含完整段落结构的节点列表
     * @deprecated 使用 extractParagraphContentBetweenBookmarks 替代
     */
    private static List<Node> extractParagraphNodesBetweenBookmarks(XWPFDocument document, 
                                                                                Node bookmarkStartNode, 
                                                                                Node bookmarkEndNode) {
        // 为了向后兼容，将新的段落内容转换为旧的格式
        List<ParagraphContent> paragraphContents = extractParagraphContentBetweenBookmarks(document, bookmarkStartNode, bookmarkEndNode);
        List<Node> allRunNodes = new ArrayList<>();
        
        for (ParagraphContent content : paragraphContents) {
            allRunNodes.addAll(content.getRunNodes());
        }
        
        return allRunNodes;
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
            // 如果无法访问书签，返回false
            return false;
        }
        return false;
    }
    
    
    /**
     * 在指定书签之前插入新书签（改进版本，保持原有书签位置不变）
     * 支持多段落书签：如果目标书签跨多个段落，则创建相同数量的段落
     */
    private static void insertBookmarkBeforeTargetBookmark(XWPFDocument document, String targetBookmarkName, String newBookmarkName) {
        // 首先检查目标书签是否为多段落书签
        BookmarkRange targetRange = findBookmarkRange(document, targetBookmarkName);
        if (targetRange.isNotFound()) {
            throw new IllegalArgumentException("目标书签 " + targetBookmarkName + " 未找到");
        }
        
        // 根据书签类型选择处理方式
        if (targetRange.isMultiParagraph()) {
            // 多段落书签：创建匹配的多段落书签
            System.out.println("🔄 检测到多段落书签，使用多段落插入方式");
            insertMultiParagraphBookmarkBefore(document, targetBookmarkName, newBookmarkName, targetRange);
        } else {
            // 单段落书签：使用原有的单段落插入方式
            System.out.println("🔄 检测到单段落书签，使用单段落插入方式");
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (int i = 0; i < paragraphs.size(); i++) {
                XWPFParagraph paragraph = paragraphs.get(i);
                if (containsBookmark(paragraph, targetBookmarkName)) {
                    insertParagraphBeforeTarget(document, paragraph, newBookmarkName);
                    break;
                }
            }
        }
    }
    
    /**
     * 在多段落书签之前插入匹配的多段落书签
     * 创建与目标书签相同段落数量的新书签结构
     */
    private static void insertMultiParagraphBookmarkBefore(XWPFDocument document, String targetBookmarkName, 
                                                          String newBookmarkName, BookmarkRange targetRange) {
        try {
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            int startIndex = targetRange.getStart();
            int endIndex = targetRange.getEnd();
            int paragraphCount = endIndex - startIndex + 1;
            
            System.out.println("📝 创建多段落书签，段落数: " + paragraphCount + 
                             " (从段落 " + startIndex + " 到 " + endIndex + ")");
            
            // 获取目标书签的第一个段落
            XWPFParagraph firstTargetParagraph = paragraphs.get(startIndex);
            CTP firstTargetCTP = firstTargetParagraph.getCTP();
            
            // 生成唯一的书签ID
            BigInteger bookmarkId = BigInteger.valueOf(System.currentTimeMillis() % 1000000);
            
            // 创建新段落列表
            List<XWPFParagraph> newParagraphs = new ArrayList<>();
            
            // 创建与目标书签相同数量的段落
            for (int i = 0; i < paragraphCount; i++) {
                XWPFParagraph newParagraph = document.createParagraph();
                
                // 复制对应目标段落的样式
                XWPFParagraph targetParagraph = paragraphs.get(startIndex + i);
                copyParagraphStyle(targetParagraph, newParagraph);
                
                // 添加初始内容
                XWPFRun run = newParagraph.createRun();
                run.setText("initialString");
                
                newParagraphs.add(newParagraph);
            }
            
            // 在段落插入到DOM之前就创建书签，避免orphaned问题
            createMultiParagraphBookmarkBeforeInsertion(newParagraphs, newBookmarkName, bookmarkId);
            
            // 将新段落插入到文档中（从前往后插入，保持顺序）
            for (int i = 0; i < newParagraphs.size(); i++) {
                XWPFParagraph newParagraph = newParagraphs.get(i);
                CTP newCTP = newParagraph.getCTP();
                
                // 在第一个目标段落之前插入
                firstTargetCTP.getDomNode().getParentNode().insertBefore(
                    newCTP.getDomNode(), firstTargetCTP.getDomNode());
            }
            
            System.out.println("✅ 多段落书签创建完成: " + newBookmarkName + 
                             " (段落数: " + paragraphCount + ")");
            
        } catch (Exception e) {
            throw new IllegalStateException("创建多段落书签失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 在段落插入到DOM之前创建多段落书签，完全避免orphaned问题
     */
    private static void createMultiParagraphBookmarkBeforeInsertion(List<XWPFParagraph> newParagraphs, 
                                                                   String bookmarkName, 
                                                                   BigInteger bookmarkId) {
        try {
            if (newParagraphs.isEmpty()) {
                return;
            }
            
            XWPFParagraph firstParagraph = newParagraphs.get(0);
            XWPFParagraph lastParagraph = newParagraphs.get(newParagraphs.size() - 1);
            
            // 在第一个段落中创建bookmarkStart
            CTP firstCTP = firstParagraph.getCTP();
            CTBookmark bookmarkStart = firstCTP.addNewBookmarkStart();
            bookmarkStart.setName(bookmarkName);
            bookmarkStart.setId(bookmarkId);
            
            // 在最后一个段落中创建bookmarkEnd
            CTP lastCTP = lastParagraph.getCTP();
            CTMarkupRange bookmarkEnd = lastCTP.addNewBookmarkEnd();
            bookmarkEnd.setId(bookmarkId);
            
            // 使用DOM操作移动bookmarkStart到第一个Run之前
            Node bookmarkStartNode = bookmarkStart.getDomNode();
            Node firstRunNode = null;
            
            // 查找第一个<w:r>节点
            NodeList children = firstCTP.getDomNode().getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                Node child = children.item(i);
                if (child.getNodeType() == Node.ELEMENT_NODE) {
                    String localName = child.getLocalName();
                    if ("r".equals(localName)) {
                        firstRunNode = child;
                        break;
                    }
                }
            }
            
            if (firstRunNode != null) {
                // 将bookmarkStart移动到第一个Run之前
                firstCTP.getDomNode().insertBefore(bookmarkStartNode, firstRunNode);
            }
            
            System.out.println("✅ 多段落书签预创建完成: " + bookmarkName + " (ID: " + bookmarkId + ")");
            
        } catch (Exception e) {
            System.err.println("❌ 预创建多段落书签失败: " + e.getMessage());
            e.printStackTrace();
            throw new IllegalStateException("预创建多段落书签失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 完全基于DOM操作创建多段落书签，避免orphaned问题
     */
    private static void createMultiParagraphBookmarkSafely(XWPFParagraph firstParagraph, 
                                                          XWPFParagraph lastParagraph, 
                                                          String bookmarkName, 
                                                          BigInteger bookmarkId) {
        try {
            // 直接使用DOM操作创建书签，完全避免Apache POI的orphaned问题
            org.w3c.dom.Document doc = firstParagraph.getDocument().getDocument().getDomNode().getOwnerDocument();
            Element firstCTPElement = (Element) firstParagraph.getCTP().getDomNode();
            Element lastCTPElement = (Element) lastParagraph.getCTP().getDomNode();
            
            // 创建bookmarkStart元素
            Element bookmarkStart = doc.createElement("w:bookmarkStart");
            bookmarkStart.setAttribute("w:name", bookmarkName);
            bookmarkStart.setAttribute("w:id", bookmarkId.toString());
            
            // 创建bookmarkEnd元素
            Element bookmarkEnd = doc.createElement("w:bookmarkEnd");
            bookmarkEnd.setAttribute("w:id", bookmarkId.toString());
            
            // 查找第一个<w:r>节点
            Node firstRunNode = null;
            NodeList children = firstCTPElement.getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                Node child = children.item(i);
                if (child.getNodeType() == Node.ELEMENT_NODE) {
                    String localName = child.getLocalName();
                    if ("r".equals(localName)) {
                        firstRunNode = child;
                        break;
                    }
                }
            }
            
            // 插入bookmarkStart到第一个Run之前
            if (firstRunNode != null) {
                firstCTPElement.insertBefore(bookmarkStart, firstRunNode);
            } else {
                // 如果没有找到Run节点，添加到段落末尾
                firstCTPElement.appendChild(bookmarkStart);
            }
            
            // 添加bookmarkEnd到最后一个段落
            lastCTPElement.appendChild(bookmarkEnd);
            
            System.out.println("✅ 多段落书签DOM创建完成: " + bookmarkName + " (ID: " + bookmarkId + ")");
            
        } catch (Exception e) {
            System.err.println("❌ 创建多段落书签失败: " + e.getMessage());
            e.printStackTrace();
            throw new IllegalStateException("创建多段落书签失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 使用DOM操作创建多段落书签，避免orphaned问题
     */
    private static void createMultiParagraphBookmarkWithDOM(XWPFParagraph firstParagraph, 
                                                           XWPFParagraph lastParagraph, 
                                                           String bookmarkName) {
        try {
            // 生成唯一的书签ID
            BigInteger bookmarkId = BigInteger.valueOf(System.currentTimeMillis() % 1000000);
            
            // 直接使用DOM操作创建书签，避免orphaned问题
            org.w3c.dom.Document doc = firstParagraph.getDocument().getDocument().getDomNode().getOwnerDocument();
            Element firstCTPElement = (Element) firstParagraph.getCTP().getDomNode();
            Element lastCTPElement = (Element) lastParagraph.getCTP().getDomNode();
            
            // 创建bookmarkStart元素
            Element bookmarkStart = doc.createElement("w:bookmarkStart");
            bookmarkStart.setAttribute("w:name", bookmarkName);
            bookmarkStart.setAttribute("w:id", bookmarkId.toString());
            
            // 创建bookmarkEnd元素
            Element bookmarkEnd = doc.createElement("w:bookmarkEnd");
            bookmarkEnd.setAttribute("w:id", bookmarkId.toString());
            
            // 查找第一个<w:r>节点
            Node firstRunNode = null;
            NodeList children = firstCTPElement.getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                Node child = children.item(i);
                if (child.getNodeType() == Node.ELEMENT_NODE) {
                    String localName = child.getLocalName();
                    if ("r".equals(localName)) {
                        firstRunNode = child;
                        break;
                    }
                }
            }
            
            // 插入bookmarkStart到第一个Run之前
            if (firstRunNode != null) {
                firstCTPElement.insertBefore(bookmarkStart, firstRunNode);
            } else {
                // 如果没有找到Run节点，添加到段落末尾
                firstCTPElement.appendChild(bookmarkStart);
            }
            
            // 添加bookmarkEnd到最后一个段落
            lastCTPElement.appendChild(bookmarkEnd);
            
            System.out.println("✅ 多段落书签DOM创建完成: " + bookmarkName + " (ID: " + bookmarkId + ")");
            
        } catch (Exception e) {
            System.err.println("❌ 创建多段落书签DOM失败: " + e.getMessage());
            e.printStackTrace();
            throw new IllegalStateException("创建多段落书签DOM失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 在段落末尾添加bookmarkEnd标记
     */
    private static void addBookmarkEndToParagraph(XWPFParagraph paragraph, String bookmarkName) {
        try {
            CTP ctp = paragraph.getCTP();
            
            // 首先查找书签ID（从第一个段落中获取）
            BigInteger bookmarkId = null;
            List<XWPFParagraph> allParagraphs = paragraph.getDocument().getParagraphs();
            for (XWPFParagraph p : allParagraphs) {
                bookmarkId = getBookmarkId(p, bookmarkName);
                if (bookmarkId != null) {
                    break;
                }
            }
            
            if (bookmarkId == null) {
                System.err.println("无法找到书签ID: " + bookmarkName);
                return;
            }
            
            // 创建bookmarkEnd标记
            CTMarkupRange bookmarkEnd = ctp.addNewBookmarkEnd();
            bookmarkEnd.setId(bookmarkId);
            
            System.out.println("✅ bookmarkEnd已添加到段落末尾，书签: " + bookmarkName + ", ID: " + bookmarkId);
            
        } catch (Exception e) {
            System.err.println("添加bookmarkEnd失败: " + e.getMessage());
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
            
            // 只添加initialString的内容，不手动添加序号（让Word编号样式自动处理）
            XWPFRun spaceRun = newParagraph.createRun();
            spaceRun.setText("initialString"); // 4个initialString
            
            // 在新段落中创建书签（包围initialString内容）
            createParagraphBookmark(newParagraph, bookmarkName);
            
            // 获取目标段落的XML节点
            CTP targetCTP = targetParagraph.getCTP();
            
            // 获取新段落的XML节点
            CTP newCTP = newParagraph.getCTP();
            
            // 在目标段落之前插入新段落
            // 使用DOM操作将新段落插入到目标段落之前
            targetCTP.getDomNode().getParentNode().insertBefore(
                newCTP.getDomNode(), targetCTP.getDomNode());
                
            System.out.println("✅ 新段落已插入，书签: " + bookmarkName);
                
        } catch (Exception e) {
            throw new IllegalStateException("在目标段落之前插入失败: " + e.getMessage(), e);
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
                
                // 复制段落样式ID（如标题1、标题2等）
                if (sourceCTP.getPPr().getPStyle() != null) {
                    targetCTP.getPPr().addNewPStyle().setVal(sourceCTP.getPPr().getPStyle().getVal());
                }
                
                // 复制其他段落属性（如对齐方式、间距等）
                if (sourceCTP.getPPr().getJc() != null) {
                    targetCTP.getPPr().setJc(sourceCTP.getPPr().getJc());
                }
                
                if (sourceCTP.getPPr().getSpacing() != null) {
                    targetCTP.getPPr().setSpacing(sourceCTP.getPPr().getSpacing());
                }
                
                // 复制缩进属性
                if (sourceCTP.getPPr().getInd() != null) {
                    targetCTP.getPPr().setInd(sourceCTP.getPPr().getInd());
                }
            } else {
                // 如果源段落没有编号样式，为目标段落设置默认编号样式
                setParagraphNumberingStyle(targetParagraph);
            }
            
        } catch (Exception e) {
            throw new IllegalStateException("复制段落样式失败: " + e.getMessage(), e);
        }
    }
    
    
    
    
    
    /**
     * 设置段落的编号样式
     */
    private static void setParagraphNumberingStyle(XWPFParagraph paragraph) {
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
            System.err.println("设置编号样式失败: " + e.getMessage());
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
            Node bookmarkStartNode = bookmarkStart.getDomNode();
            Node firstRunNode = null;
            
            // 查找第一个<w:r>节点
            NodeList children = ctp.getDomNode().getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                Node child = children.item(i);
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
            e.printStackTrace();
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
            Node paragraphNode = ctp.getDomNode();
            
            // 查找bookmarkStart节点
            Node bookmarkStartNode = findBookmarkStartNode(paragraphNode, bookmarkId);
            if (bookmarkStartNode == null) {
                System.err.println("未找到bookmarkStart节点，ID: " + bookmarkId);
                return "";
            }
            
            // 查找对应的bookmarkEnd节点（可能在当前段落或后续段落中）
            Node bookmarkEndNode = findBookmarkEndNodeInDocument(paragraph, bookmarkId);
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
    private static Node findBookmarkStartNode(Node paragraphNode, BigInteger bookmarkId) {
        NodeList children = paragraphNode.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            if (child.getLocalName() != null && child.getLocalName().equals("bookmarkStart")) {
                // 检查ID是否匹配
                NamedNodeMap attributes = child.getAttributes();
                if (attributes != null) {
                    Node idAttr = attributes.getNamedItem("w:id");
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
    private static Node findBookmarkEndNode(Node paragraphNode, BigInteger bookmarkId) {
        NodeList children = paragraphNode.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            
            // 打印所有子节点信息用于调试
            System.out.println("🔍 检查子节点: " + child.getNodeName() + ", 本地名: " + child.getLocalName() + ", 类型: " + child.getNodeType());
            
            if (child.getLocalName() != null && child.getLocalName().equals("bookmarkEnd")) {
                System.out.println("🎯 找到bookmarkEnd节点！");
                // 检查ID是否匹配
                NamedNodeMap attributes = child.getAttributes();
                if (attributes != null) {
                    // 打印所有属性
                    for (int j = 0; j < attributes.getLength(); j++) {
                        Node attr = attributes.item(j);
                        System.out.println("   属性: " + attr.getNodeName() + " = " + attr.getNodeValue());
                    }
                    
                    Node idAttr = attributes.getNamedItem("w:id");
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
    private static Node findBookmarkEndNodeInDocument(XWPFParagraph startParagraph, BigInteger bookmarkId) {
        try {
            System.out.println("🔍 查找bookmarkEnd节点，ID: " + bookmarkId);
            
            // 首先在当前段落中查找
            CTP ctp = startParagraph.getCTP();
            Node paragraphNode = ctp.getDomNode();
            Node bookmarkEndNode = findBookmarkEndNode(paragraphNode, bookmarkId);
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
                Node paragraphNode2 = paragraphCTP.getDomNode();
                
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
                Node paragraphNode2 = paragraphCTP.getDomNode();
                
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
                Node documentNode = documentCT.getDomNode();
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
    private static Node findBookmarkEndNodeInDocumentBody(Node documentNode, BigInteger bookmarkId) {
        NodeList children = documentNode.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node child = children.item(i);
            System.out.println("🔍 检查文档主体子节点: " + child.getNodeName() + ", 本地名: " + child.getLocalName());
            
            if (child.getLocalName() != null && child.getLocalName().equals("bookmarkEnd")) {
                System.out.println("🎯 在文档主体中找到bookmarkEnd节点！");
                NamedNodeMap attributes = child.getAttributes();
                if (attributes != null) {
                    Node idAttr = attributes.getNamedItem("w:id");
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
                Node result = findBookmarkEndNodeInDocumentBody(child, bookmarkId);
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
    private static String extractTextBetweenNodes(Node startNode, Node endNode) {
        StringBuilder content = new StringBuilder();
        
        // 如果startNode和endNode在同一个段落中
        if (startNode.getParentNode().equals(endNode.getParentNode())) {
            // 从startNode的下一个兄弟节点开始，到endNode的前一个兄弟节点结束
            Node current = startNode.getNextSibling();
            while (current != null && !current.equals(endNode)) {
                if (current.getNodeType() == Node.TEXT_NODE) {
                    content.append(current.getNodeValue());
                } else if (current.getNodeType() == Node.ELEMENT_NODE) {
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
            Node current = startNode.getNextSibling();
            while (current != null) {
                if (current.getNodeType() == Node.TEXT_NODE) {
                    content.append(current.getNodeValue());
                } else if (current.getNodeType() == Node.ELEMENT_NODE) {
                    String text = extractTextFromElement(current);
                    if (!text.isEmpty()) {
                        content.append(text);
                    }
                }
                current = current.getNextSibling();
            }
            
            // 然后提取中间段落的完整内容
            Node startParent = startNode.getParentNode();
            Node endParent = endNode.getParentNode();
            Node currentParent = startParent.getNextSibling();
            
            while (currentParent != null && !currentParent.equals(endParent)) {
                if (currentParent.getNodeType() == Node.ELEMENT_NODE) {
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
                if (current.getNodeType() == Node.TEXT_NODE) {
                    content.append(current.getNodeValue());
                } else if (current.getNodeType() == Node.ELEMENT_NODE) {
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
    private static String extractTextFromElement(Node element) {
        StringBuilder text = new StringBuilder();
        
        if (element.getNodeType() == Node.TEXT_NODE) {
            text.append(element.getNodeValue());
        } else if (element.getNodeType() == Node.ELEMENT_NODE) {
            // 递归提取子节点的文本
            NodeList children = element.getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                text.append(extractTextFromElement(children.item(i)));
            }
        }
        
        return text.toString();
    }
    
    /**
     * 提取书签之间的run节点（包含格式信息）
     * 修复：支持多段落书签，提取实际的XML run节点而不是纯文本，以保持所有格式
     */
    private static List<Node> extractRunNodesBetweenBookmarks(XWPFParagraph paragraph, BigInteger bookmarkId) {
        List<Node> runNodes = new ArrayList<>();
        
        try {
            CTP ctp = paragraph.getCTP();
            Node paragraphNode = ctp.getDomNode();
            
            // 查找bookmarkStart节点
            Node bookmarkStartNode = findBookmarkStartNode(paragraphNode, bookmarkId);
            if (bookmarkStartNode == null) {
                System.err.println("未找到bookmarkStart节点，ID: " + bookmarkId);
                return runNodes;
            }
            
            // 查找对应的bookmarkEnd节点（支持跨段落）
            Node bookmarkEndNode = findBookmarkEndNodeInDocument(paragraph, bookmarkId);
            if (bookmarkEndNode == null) {
                System.err.println("未找到bookmarkEnd节点，ID: " + bookmarkId);
                return runNodes;
            }
            
            // 使用新的多段落支持方法提取节点
            XWPFDocument document = paragraph.getDocument();
            runNodes = extractParagraphNodesBetweenBookmarks(document, bookmarkStartNode, bookmarkEndNode);
            
            System.out.println("✅ 提取到 " + runNodes.size() + " 个run节点，支持多段落书签，保持格式信息");
            
        } catch (Exception e) {
            System.err.println("提取run节点失败: " + e.getMessage());
        }
        
        return runNodes;
    }
    
    
    /**
     * 替换书签之间的内容，同时保持书签标记不变
     * 使用DOM操作精确替换内容，避免破坏书签结构
     */
    private static void replaceContentBetweenBookmarks(XWPFParagraph paragraph, BigInteger bookmarkId, String newContent) {
        try {
            CTP ctp = paragraph.getCTP();
            Node paragraphNode = ctp.getDomNode();
            
            // 查找bookmarkStart和bookmarkEnd节点
            Node bookmarkStartNode = findBookmarkStartNode(paragraphNode, bookmarkId);
            Node bookmarkEndNode = findBookmarkEndNode(paragraphNode, bookmarkId);
            
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
     * 替换书签之间的内容为run节点（保持格式）
     * 修复：支持多段落书签，使用run节点替换内容以保持所有格式信息
     */
    private static void replaceContentBetweenBookmarksWithRunNodes(XWPFParagraph paragraph, BigInteger bookmarkId, List<Node> runNodes) {
        try {
            CTP ctp = paragraph.getCTP();
            Node paragraphNode = ctp.getDomNode();
            
            // 查找bookmarkStart节点
            Node bookmarkStartNode = findBookmarkStartNode(paragraphNode, bookmarkId);
            if (bookmarkStartNode == null) {
                System.err.println("无法找到bookmarkStart节点，ID: " + bookmarkId);
                return;
            }
            
            // 查找bookmarkEnd节点（支持跨段落）
            Node bookmarkEndNode = findBookmarkEndNodeInDocument(paragraph, bookmarkId);
            if (bookmarkEndNode == null) {
                System.err.println("无法找到bookmarkEnd节点，ID: " + bookmarkId);
                return;
            }
            
            // 获取文档对象以支持多段落操作
            XWPFDocument document = paragraph.getDocument();
            
            // 删除bookmarkStart和bookmarkEnd之间的所有内容节点（支持多段落）
            removeContentBetweenBookmarksMultiParagraph(document, bookmarkStartNode, bookmarkEndNode);
            
            // 在bookmarkStart之后插入节点（支持多段落）
            insertParagraphNodesAfterBookmarkStart(document, bookmarkStartNode, runNodes);
            
            System.out.println("✅ 书签内容已替换为run节点，支持多段落，保持格式，ID: " + bookmarkId);
            
        } catch (Exception e) {
            System.err.println("替换书签内容为run节点失败: " + e.getMessage());
        }
    }
    
    /**
     * 删除两个书签标记之间的所有内容节点
     */
    private static void removeContentBetweenBookmarks(Node bookmarkStartNode, Node bookmarkEndNode) {
        Node current = bookmarkStartNode.getNextSibling();
        while (current != null && !current.equals(bookmarkEndNode)) {
            Node next = current.getNextSibling();
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
     * 删除多段落书签之间的所有内容节点
     * 支持跨段落的书签内容删除
     */
    private static void removeContentBetweenBookmarksMultiParagraph(XWPFDocument document, 
                                                                   Node bookmarkStartNode, 
                                                                   Node bookmarkEndNode) {
        try {
            // 如果bookmarkStart和bookmarkEnd在同一个段落中
            if (bookmarkStartNode.getParentNode().equals(bookmarkEndNode.getParentNode())) {
                // 单段落情况：使用原有逻辑
                removeContentBetweenBookmarks(bookmarkStartNode, bookmarkEndNode);
            } else {
                // 多段落情况：需要删除中间段落和部分段落内容
                Node startParent = bookmarkStartNode.getParentNode();
                Node endParent = bookmarkEndNode.getParentNode();
                
                // 获取段落索引
                int startParagraphIndex = findParagraphIndexContainingNode(document, startParent);
                int endParagraphIndex = findParagraphIndexContainingNode(document, endParent);
                
                if (startParagraphIndex != -1 && endParagraphIndex != -1) {
                    List<XWPFParagraph> paragraphs = document.getParagraphs();
                    
                    // 删除起始段落中bookmarkStart之后的内容
                    Node current = bookmarkStartNode.getNextSibling();
                    while (current != null) {
                        Node next = current.getNextSibling();
                        if (current.getLocalName() != null && 
                            !current.getLocalName().equals("bookmarkStart") && 
                            !current.getLocalName().equals("bookmarkEnd")) {
                            current.getParentNode().removeChild(current);
                        }
                        current = next;
                    }
                    
                    // 删除中间段落（如果存在）
                    for (int i = startParagraphIndex + 1; i < endParagraphIndex; i++) {
                        XWPFParagraph paragraph = paragraphs.get(i);
                        CTP ctp = paragraph.getCTP();
                        Node paragraphNode = ctp.getDomNode();
                        
                        // 删除段落中的所有内容，但保留段落结构
                        NodeList children = paragraphNode.getChildNodes();
                        List<Node> nodesToRemove = new ArrayList<>();
                        for (int j = 0; j < children.getLength(); j++) {
                            Node child = children.item(j);
                            if (child.getLocalName() != null && 
                                !child.getLocalName().equals("pPr")) { // 保留段落属性
                                nodesToRemove.add(child);
                            }
                        }
                        
                        for (Node node : nodesToRemove) {
                            paragraphNode.removeChild(node);
                        }
                    }
                    
                    // 删除结束段落中bookmarkEnd之前的内容
                    if (startParagraphIndex != endParagraphIndex) {
                        Node endCurrent = endParent.getFirstChild();
                        while (endCurrent != null && !endCurrent.equals(bookmarkEndNode)) {
                            Node next = endCurrent.getNextSibling();
                            if (endCurrent.getLocalName() != null && 
                                !endCurrent.getLocalName().equals("bookmarkStart") && 
                                !endCurrent.getLocalName().equals("bookmarkEnd")) {
                                endCurrent.getParentNode().removeChild(endCurrent);
                            }
                            endCurrent = next;
                        }
                    }
                }
            }
            
            System.out.println("✅ 多段落书签内容删除完成");
            
        } catch (Exception e) {
            System.err.println("删除多段落书签内容失败: " + e.getMessage());
        }
    }
    
    /**
     * 在bookmarkStart之后插入新内容
     */
    private static void insertContentAfterBookmarkStart(XWPFParagraph paragraph, Node bookmarkStartNode, String newContent) {
        try {
            // 创建新的run来包含内容
            XWPFRun newRun = paragraph.createRun();
            newRun.setText(newContent);
            
            // 获取新run的DOM节点
            Node newRunNode = newRun.getCTR().getDomNode();
            
            // 将新run插入到bookmarkStart之后
            bookmarkStartNode.getParentNode().insertBefore(newRunNode, bookmarkStartNode.getNextSibling());
            
        } catch (Exception e) {
            System.err.println("插入新内容失败: " + e.getMessage());
        }
    }
    
    /**
     * 在bookmarkStart之后插入run节点（保持格式）
     * 修复：克隆run节点以保持所有格式信息，并保持正确的顺序
     */
    private static void insertRunNodesAfterBookmarkStart(XWPFParagraph paragraph, Node bookmarkStartNode, List<Node> runNodes) {
        try {
            org.w3c.dom.Document ownerDocument = bookmarkStartNode.getOwnerDocument();
            Node parentNode = bookmarkStartNode.getParentNode();
            Node insertAfterNode = bookmarkStartNode;
            
            for (Node runNode : runNodes) {
                // 深度克隆run节点以保持所有格式属性
                Node clonedRunNode = runNode.cloneNode(true);
                
                // 如果节点来自不同的文档，需要导入到当前文档
                if (!ownerDocument.equals(runNode.getOwnerDocument())) {
                    clonedRunNode = ownerDocument.importNode(clonedRunNode, true);
                }
                
                // 将克隆的run节点插入到正确的位置，保持顺序
                if (insertAfterNode.getNextSibling() == null) {
                    // 如果没有下一个兄弟节点，直接追加到末尾
                    parentNode.appendChild(clonedRunNode);
                } else {
                    // 插入到insertAfterNode之后
                    parentNode.insertBefore(clonedRunNode, insertAfterNode.getNextSibling());
                }
                
                // 更新insertAfterNode为刚插入的节点，确保下一个节点插入在它之后
                insertAfterNode = clonedRunNode;
            }
            
            System.out.println("✅ 成功插入 " + runNodes.size() + " 个带格式的run节点，保持正确顺序");
            
        } catch (Exception e) {
            System.err.println("插入run节点失败: " + e.getMessage());
        }
    }
    
    /**
     * 在bookmarkStart之后插入段落节点（支持多段落书签）
     * 处理单段落和多段落内容的插入
     */
    private static void insertParagraphNodesAfterBookmarkStart(XWPFDocument document, 
                                                              Node bookmarkStartNode, 
                                                              List<Node> paragraphNodes) {
        try {
            // 检查是否是多段落内容（包含段落节点）
            boolean isMultiParagraph = paragraphNodes.stream()
                .anyMatch(node -> node.getLocalName() != null && node.getLocalName().equals("p"));
            
            if (isMultiParagraph) {
                // 多段落情况：需要插入到文档级别，而不是段落内
                insertMultiParagraphContent(document, bookmarkStartNode, paragraphNodes);
            } else {
                // 单段落情况：在段落内插入run节点
                insertRunNodesAfterBookmarkStart(
                    findParagraphContainingNode(document, bookmarkStartNode), 
                    bookmarkStartNode, 
                    paragraphNodes
                );
            }
            
            System.out.println("✅ 成功插入段落节点，支持多段落书签");
            
        } catch (Exception e) {
            System.err.println("插入段落节点失败: " + e.getMessage());
        }
    }
    
    /**
     * 插入多段落内容到文档中
     */
    private static void insertMultiParagraphContent(XWPFDocument document, 
                                                   Node bookmarkStartNode, 
                                                   List<Node> paragraphNodes) {
        try {
            // 找到bookmarkStart所在的段落
            XWPFParagraph startParagraph = findParagraphContainingNode(document, bookmarkStartNode);
            if (startParagraph == null) {
                System.err.println("无法找到bookmarkStart所在的段落");
                return;
            }
            
            // 获取文档的段落列表
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            int startParagraphIndex = -1;
            for (int i = 0; i < paragraphs.size(); i++) {
                if (paragraphs.get(i) == startParagraph) {
                    startParagraphIndex = i;
                    break;
                }
            }
            
            if (startParagraphIndex == -1) {
                System.err.println("无法确定起始段落索引");
                return;
            }
            
            // 在起始段落之后插入新的段落
            for (int i = 0; i < paragraphNodes.size(); i++) {
                Node paragraphNode = paragraphNodes.get(i);
                
                // 创建新段落
                XWPFParagraph newParagraph = document.createParagraph();
                CTP newCTP = newParagraph.getCTP();
                
                // 克隆段落节点内容到新段落
                Node clonedNode = paragraphNode.cloneNode(true);
                org.w3c.dom.Document ownerDocument = newCTP.getDomNode().getOwnerDocument();
                if (!ownerDocument.equals(clonedNode.getOwnerDocument())) {
                    clonedNode = ownerDocument.importNode(clonedNode, true);
                }
                
                // 将克隆的段落内容添加到新段落
                NodeList children = clonedNode.getChildNodes();
                for (int j = 0; j < children.getLength(); j++) {
                    Node child = children.item(j);
                    newCTP.getDomNode().appendChild(child.cloneNode(true));
                }
                
                // 将新段落插入到文档中
                CTP startCTP = startParagraph.getCTP();
                startCTP.getDomNode().getParentNode().insertBefore(
                    newCTP.getDomNode(), 
                    startCTP.getDomNode().getNextSibling()
                );
            }
            
        } catch (Exception e) {
            System.err.println("插入多段落内容失败: " + e.getMessage());
        }
    }
    
    /**
     * 查找包含指定节点的段落
     */
    private static XWPFParagraph findParagraphContainingNode(XWPFDocument document, Node targetNode) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (XWPFParagraph paragraph : paragraphs) {
            CTP ctp = paragraph.getCTP();
            Node paragraphNode = ctp.getDomNode();
            
            if (isNodeContainedIn(paragraphNode, targetNode)) {
                return paragraph;
            }
        }
        return null;
    }
    
    
    
    /**
     * 为书签设置内容并保持编号样式
     * 修复：使用DOM操作保持书签结构，避免破坏bookmarkStart和bookmarkEnd位置
     */
    private static void setBookmarkContent(XWPFDocument document, String bookmarkName, String content) {
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
                    
                    // 使用DOM操作替换内容，保持书签结构
                    replaceContentBetweenBookmarks(paragraph, bookmarkId, content);
                    
                    System.out.println("✅ 书签内容已更新，保持书签结构: " + bookmarkName);
                    break;
                    
                } catch (Exception e) {
                    throw new IllegalStateException("设置书签内容失败: " + e.getMessage(), e);
                }
            }
        }
    }
    
    /**
     * 为书签设置段落内容（支持多段落书签）
     * 保持段落结构和格式信息
     */
    private static void setBookmarkContentFromParagraphContent(XWPFDocument document, String bookmarkName, List<ParagraphContent> paragraphContents) {
        // 检查目标书签是否为多段落
        BookmarkRange targetRange = findBookmarkRange(document, bookmarkName);
        if (targetRange.isNotFound()) {
            throw new IllegalArgumentException("目标书签 " + bookmarkName + " 未找到");
        }
        
        if (targetRange.getStart() == targetRange.getEnd()) {
            // 单段落书签：将所有内容合并到一个段落
            setSingleParagraphContentFromParagraphContent(document, bookmarkName, paragraphContents);
        } else {
            // 多段落书签：按段落分布内容
            setMultiParagraphContentFromParagraphContent(document, bookmarkName, paragraphContents, targetRange);
        }
    }
    
    /**
     * 为单段落书签设置段落内容
     */
    private static void setSingleParagraphContentFromParagraphContent(XWPFDocument document, String bookmarkName, List<ParagraphContent> paragraphContents) {
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
                    
                    // 合并所有段落的run节点
                    List<Node> allRunNodes = new ArrayList<>();
                    for (ParagraphContent content : paragraphContents) {
                        allRunNodes.addAll(content.getRunNodes());
                    }
                    
                    // 使用DOM操作替换内容为run节点，保持书签结构和格式
                    replaceContentBetweenBookmarksWithRunNodes(paragraph, bookmarkId, allRunNodes);
                    
                    System.out.println("✅ 单段落书签内容已更新，保持格式和书签结构: " + bookmarkName);
                    break;
                    
                } catch (Exception e) {
                    throw new IllegalStateException("设置单段落书签内容失败: " + e.getMessage(), e);
                }
            }
        }
    }
    
    /**
     * 为多段落书签设置段落内容
     */
    private static void setMultiParagraphContentFromParagraphContent(XWPFDocument document, String bookmarkName, 
                                                                    List<ParagraphContent> paragraphContents, BookmarkRange targetRange) {
        try {
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            int startIndex = targetRange.getStart();
            int endIndex = targetRange.getEnd();
            
            System.out.println("📝 设置多段落书签内容，段落数: " + paragraphContents.size() + 
                             " (目标段落 " + startIndex + " 到 " + endIndex + ")");
            
            // 确保源段落数和目标段落数匹配
            int targetParagraphCount = endIndex - startIndex + 1;
            if (paragraphContents.size() != targetParagraphCount) {
                System.err.println("⚠️ 源段落数(" + paragraphContents.size() + 
                                 ")与目标段落数(" + targetParagraphCount + ")不匹配");
            }
            
            // 获取书签ID（只在第一个段落中查找）
            BigInteger bookmarkId = null;
            if (startIndex < paragraphs.size()) {
                bookmarkId = getBookmarkId(paragraphs.get(startIndex), bookmarkName);
            }
            
            if (bookmarkId == null) {
                throw new IllegalStateException("无法找到书签ID: " + bookmarkName);
            }
            
            // 为每个目标段落设置对应的源段落内容
            for (int i = 0; i < Math.min(paragraphContents.size(), targetParagraphCount); i++) {
                int targetParagraphIndex = startIndex + i;
                if (targetParagraphIndex < paragraphs.size()) {
                    XWPFParagraph targetParagraph = paragraphs.get(targetParagraphIndex);
                    ParagraphContent sourceContent = paragraphContents.get(i);
                    
                    if (i == 0) {
                        // 第一个段落：替换bookmarkStart和bookmarkEnd之间的内容
                        replaceContentBetweenBookmarksWithRunNodes(targetParagraph, bookmarkId, sourceContent.getRunNodes());
                    } else {
                        // 中间段落：直接替换整个段落的内容
                        replaceParagraphContentWithRunNodes(targetParagraph, sourceContent.getRunNodes());
                    }
                }
            }
            
            System.out.println("✅ 多段落书签内容已更新，保持段落结构: " + bookmarkName);
            
        } catch (Exception e) {
            throw new IllegalStateException("设置多段落书签内容失败: " + e.getMessage(), e);
        }
    }
    
    /**
     * 替换段落内容为run节点（保持格式）
     * 用于多段落书签的中间段落
     */
    private static void replaceParagraphContentWithRunNodes(XWPFParagraph paragraph, List<Node> runNodes) {
        try {
            CTP ctp = paragraph.getCTP();
            Node paragraphNode = ctp.getDomNode();
            
            // 删除段落中的所有内容节点（保留段落属性）
            List<Node> nodesToRemove = new ArrayList<>();
            for (int i = 0; i < paragraphNode.getChildNodes().getLength(); i++) {
                Node child = paragraphNode.getChildNodes().item(i);
                if (child.getNodeType() == Node.ELEMENT_NODE) {
                    String localName = child.getLocalName();
                    // 保留段落属性节点，删除其他内容节点
                    if (!"pPr".equals(localName)) {
                        nodesToRemove.add(child);
                    }
                }
            }
            
            for (Node node : nodesToRemove) {
                paragraphNode.removeChild(node);
            }
            
            // 插入新的run节点
            for (Node runNode : runNodes) {
                Node importedNode = paragraphNode.getOwnerDocument().importNode(runNode, true);
                paragraphNode.appendChild(importedNode);
            }
            
            System.out.println("✅ 段落内容已替换为run节点，保持格式");
            
        } catch (Exception e) {
            System.err.println("替换段落内容为run节点失败: " + e.getMessage());
        }
    }
    
    /**
     * 为书签设置run节点内容（保持格式）
     * 修复：使用run节点设置内容以保持所有格式信息
     */
    private static void setBookmarkContentFromRunNodes(XWPFDocument document, String bookmarkName, List<Node> runNodes) {
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
                    
                    // 使用DOM操作替换内容为run节点，保持书签结构和格式
                    replaceContentBetweenBookmarksWithRunNodes(paragraph, bookmarkId, runNodes);
                    
                    System.out.println("✅ 书签内容已更新为run节点，保持格式和书签结构: " + bookmarkName);
                    break;
                    
                } catch (Exception e) {
                    throw new IllegalStateException("设置书签run节点内容失败: " + e.getMessage(), e);
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
     * 从文件中获取书签包含的段落数量
     * @param documentPath 文档路径
     * @param bookmarkName 书签名称
     * @return 书签包含的段落数量
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XmlException
     */
    public static int getBookmarkParagraphCountFromFile(String documentPath, String bookmarkName) 
                                                       throws IOException, InvalidFormatException, XmlException {
        try (FileInputStream fis = new FileInputStream(documentPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            return getBookmarkParagraphCount(document, bookmarkName);
        }
    }
    
    /**
     * 比较两个书签中对应段落的样式是否一致
     * @param documentPath 文档路径
     * @param bookmarkName1 第一个书签名称
     * @param bookmarkName2 第二个书签名称
     * @return 样式是否一致
     * @throws IOException
     * @throws InvalidFormatException
     * @throws XmlException
     */
    public static boolean compareBookmarkParagraphStyles(String documentPath, String bookmarkName1, String bookmarkName2) 
                                                         throws IOException, InvalidFormatException, XmlException {
        try (FileInputStream fis = new FileInputStream(documentPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            return compareBookmarkParagraphStyles(document, bookmarkName1, bookmarkName2);
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
     * 检查段落是否使用编号样式
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
                } catch (Exception e) {
                    // 如果无法检查，返回false
                }
                return false;
            }
        }
        return false;
    }
    
    /**
     * 获取书签包含的段落数量
     * @param document 文档对象
     * @param bookmarkName 书签名称
     * @return 书签包含的段落数量
     */
    private static int getBookmarkParagraphCount(XWPFDocument document, String bookmarkName) {
        BookmarkRange range = findBookmarkRange(document, bookmarkName);
        if (range.isNotFound()) {
            return 0;
        }
        
        // 计算书签跨越的段落数量
        int startIndex = range.getStartParagraphIndex();
        int endIndex = range.getEndParagraphIndex();
        
        return endIndex - startIndex + 1;
    }
    
    /**
     * 比较两个书签中对应段落的样式是否一致
     * @param document 文档对象
     * @param bookmarkName1 第一个书签名称
     * @param bookmarkName2 第二个书签名称
     * @return 样式是否一致
     */
    private static boolean compareBookmarkParagraphStyles(XWPFDocument document, String bookmarkName1, String bookmarkName2) {
        BookmarkRange range1 = findBookmarkRange(document, bookmarkName1);
        BookmarkRange range2 = findBookmarkRange(document, bookmarkName2);

        System.out.println("🔍 比较书签段落样式 - " + bookmarkName1 + " vs " + bookmarkName2);
        System.out.println("📝 " + bookmarkName1 + " 范围: " + range1);
        System.out.println("📝 " + bookmarkName2 + " 范围: " + range2);

        if (range1.isNotFound() || range2.isNotFound()) {
            System.out.println("❌ 书签未找到");
            return false;
        }

        // 检查段落数量是否相同
        int count1 = range1.getEndParagraphIndex() - range1.getStartParagraphIndex() + 1;
        int count2 = range2.getEndParagraphIndex() - range2.getStartParagraphIndex() + 1;

        System.out.println("📝 " + bookmarkName1 + " 段落数: " + count1);
        System.out.println("📝 " + bookmarkName2 + " 段落数: " + count2);

        if (count1 != count2) {
            System.out.println("❌ 段落数量不同");
            return false;
        }

        // 比较每个对应段落的样式
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (int i = 0; i < count1; i++) {
            int index1 = range1.getStartParagraphIndex() + i;
            int index2 = range2.getStartParagraphIndex() + i;
            XWPFParagraph para1 = paragraphs.get(index1);
            XWPFParagraph para2 = paragraphs.get(index2);

            System.out.println("📋 比较第 " + i + " 个段落: 索引 " + index1 + " vs " + index2);

            if (!compareParagraphStyles(para1, para2)) {
                System.out.println("❌ 第 " + i + " 个段落样式不同");
                return false;
            }

            System.out.println("✅ 第 " + i + " 个段落样式相同");
        }

        System.out.println("✅ 所有段落样式都相同");
        return true;
    }
    
    /**
     * 比较两个段落的样式是否一致
     * @param para1 第一个段落
     * @param para2 第二个段落
     * @return 样式是否一致
     */
    private static boolean compareParagraphStyles(XWPFParagraph para1, XWPFParagraph para2) {
        // 获取段落内容用于日志输出
        String content1 = getParagraphText(para1);
        String content2 = getParagraphText(para2);
        
        // 比较段落对齐方式
        if (para1.getAlignment() != para2.getAlignment()) {
            System.out.println("❌ 段落对齐方式不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1对齐: " + para1.getAlignment());
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2对齐: " + para2.getAlignment());
            return false;
        }

        // 比较段落间距
        if (para1.getSpacingBefore() != para2.getSpacingBefore()) {
            System.out.println("❌ 段前间距不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1段前间距: " + para1.getSpacingBefore());
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2段前间距: " + para2.getSpacingBefore());
            return false;
        }
        if (para1.getSpacingAfter() != para2.getSpacingAfter()) {
            System.out.println("❌ 段后间距不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1段后间距: " + para1.getSpacingAfter());
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2段后间距: " + para2.getSpacingAfter());
            return false;
        }
        if (para1.getSpacingBetween() != para2.getSpacingBetween()) {
            System.out.println("❌ 行间距不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1行间距: " + para1.getSpacingBetween());
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2行间距: " + para2.getSpacingBetween());
            return false;
        }

        // 比较段落缩进
        if (para1.getIndentationLeft() != para2.getIndentationLeft()) {
            System.out.println("❌ 左缩进不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1左缩进: " + para1.getIndentationLeft());
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2左缩进: " + para2.getIndentationLeft());
            return false;
        }
        if (para1.getIndentationRight() != para2.getIndentationRight()) {
            System.out.println("❌ 右缩进不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1右缩进: " + para1.getIndentationRight());
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2右缩进: " + para2.getIndentationRight());
            return false;
        }
        if (para1.getIndentationFirstLine() != para2.getIndentationFirstLine()) {
            System.out.println("❌ 首行缩进不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1首行缩进: " + para1.getIndentationFirstLine());
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2首行缩进: " + para2.getIndentationFirstLine());
            return false;
        }
        if (para1.getIndentationHanging() != para2.getIndentationHanging()) {
            System.out.println("❌ 悬挂缩进不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1悬挂缩进: " + para1.getIndentationHanging());
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2悬挂缩进: " + para2.getIndentationHanging());
            return false;
        }

        // 比较编号样式
        BigInteger numId1 = para1.getNumID();
        BigInteger numId2 = para2.getNumID();
        if ((numId1 == null && numId2 != null) || (numId1 != null && !numId1.equals(numId2))) {
            System.out.println("❌ 编号ID不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1编号ID: " + numId1);
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2编号ID: " + numId2);
            return false;
        }
        
        BigInteger numIlvl1 = para1.getNumIlvl();
        BigInteger numIlvl2 = para2.getNumIlvl();
        if ((numIlvl1 == null && numIlvl2 != null) || (numIlvl1 != null && !numIlvl1.equals(numIlvl2))) {
            System.out.println("❌ 编号级别不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1编号级别: " + numIlvl1);
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2编号级别: " + numIlvl2);
            return false;
        }

        // 比较段落样式ID
        String style1 = para1.getStyle();
        String style2 = para2.getStyle();
        if ((style1 == null && style2 != null) || (style1 != null && !style1.equals(style2))) {
            System.out.println("❌ 段落样式ID不同:");
            System.out.println("   段落1内容: \"" + content1 + "\"");
            System.out.println("   段落1样式ID: '" + style1 + "'");
            System.out.println("   段落2内容: \"" + content2 + "\"");
            System.out.println("   段落2样式ID: '" + style2 + "'");
            return false;
        }

        return true;
    }
    
    /**
     * 获取段落的文本内容
     * @param paragraph 段落对象
     * @return 段落文本内容
     */
    private static String getParagraphText(XWPFParagraph paragraph) {
        if (paragraph == null) {
            return "";
        }
        String text = paragraph.getText();
        return text != null ? text.trim() : "";
    }

}
