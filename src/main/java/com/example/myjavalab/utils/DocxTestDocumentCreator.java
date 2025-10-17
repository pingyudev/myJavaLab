package com.example.myjavalab.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

/**
 * 创建包含测试书签的DOCX文档
 */
public class DocxTestDocumentCreator {

    public static void createTestDocument(String outputPath) throws IOException {
        XWPFDocument document = new XWPFDocument();
        
        // 创建标题
        XWPFParagraph titleParagraph = document.createParagraph();
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText("追求硕士学位的原因");
        titleRun.setBold(true);
        titleRun.setFontSize(16);
        
        // 创建空行
        document.createParagraph();
        
        // 创建第一个原因段落
        XWPFParagraph reason1Paragraph = document.createParagraph();
        XWPFRun reason1Run = reason1Paragraph.createRun();
        reason1Run.setText("1. 提升解决复杂问题的能力。");
        
        // 创建包含labelA书签的段落（第二个原因）
        XWPFParagraph bookmarkParagraph = document.createParagraph();
        System.out.println("🔧 创建第2个段落，段落索引: " + (document.getParagraphs().size() - 1));
        
        // 添加序号
        XWPFRun numberRun = bookmarkParagraph.createRun();
        numberRun.setText("2. ");
        
        // 添加粗体标题部分
        XWPFRun boldRun = bookmarkParagraph.createRun();
        boldRun.setText("提升职场竞争力，拥抱AI浪潮：");
        boldRun.setBold(true);
        
        // 添加详细内容
        XWPFRun contentRun = bookmarkParagraph.createRun();
        contentRun.setText(" 当前大型科技公司偏好高学历人才，硕士学位将显著提升我的职场竞争力。更关键的是，AI浪潮汹涌，我需要抓住机遇，系统更新并掌握AI技术栈，以应对未来职场对AI人才的迫切需求。");
        
        System.out.println("🔧 段落内容: '" + bookmarkParagraph.getText() + "'");
        System.out.println("🔧 准备创建书签 labelA...");
        
        // 在段落结束处创建书签（包围整个段落内容）
        createBookmark(bookmarkParagraph, "labelA");
        
        // 立即检查书签位置
        System.out.println("🔍 创建书签后立即检查:");
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph p = paragraphs.get(i);
            try {
                CTP ctp = p.getCTP();
                CTBookmark[] bookmarks = ctp.getBookmarkStartArray();
                if (bookmarks != null && bookmarks.length > 0) {
                    for (CTBookmark bookmark : bookmarks) {
                        if ("labelA".equals(bookmark.getName())) {
                            System.out.println("  📌 labelA 在段落 " + i + ": '" + p.getText() + "'");
                        }
                    }
                }
            } catch (Exception e) {
                // 忽略错误
            }
        }
        
        // 创建另一个段落
        XWPFParagraph anotherParagraph = document.createParagraph();
        XWPFRun anotherRun = anotherParagraph.createRun();
        anotherRun.setText("这是另一个段落，用于测试文档结构。");
        
        // 保存文档
        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            document.write(fos);
        }
        
        document.close();
        System.out.println("测试文档已创建: " + outputPath);
    }
    
    /**
     * 在段落开始处创建书签开始标记
     */
    private static void createBookmarkStart(XWPFParagraph paragraph, String bookmarkName) {
        try {
            // 获取段落的底层XML对象
            CTP ctp = paragraph.getCTP();
            
            // 在段落开始处创建书签开始标记
            CTBookmark bookmarkStart = ctp.addNewBookmarkStart();
            bookmarkStart.setName(bookmarkName);
            bookmarkStart.setId(BigInteger.valueOf(0));
            
        } catch (Exception e) {
            System.err.println("创建书签开始标记失败: " + e.getMessage());
        }
    }
    
    /**
     * 在段落结束处创建书签结束标记
     */
    private static void createBookmarkEnd(XWPFParagraph paragraph, String bookmarkName) {
        try {
            // 获取段落的底层XML对象
            CTP ctp = paragraph.getCTP();
            
            // 在段落结束处创建书签结束标记
            CTMarkupRange bookmarkEnd = ctp.addNewBookmarkEnd();
            bookmarkEnd.setId(BigInteger.valueOf(0));
            
        } catch (Exception e) {
            System.err.println("创建书签结束标记失败: " + e.getMessage());
        }
    }
    
    /**
     * 在段落中创建真正的Word书签（包围整个段落内容）
     */
    private static void createBookmark(XWPFParagraph paragraph, String bookmarkName) {
        try {
            // 获取段落的底层XML对象
            CTP ctp = paragraph.getCTP();
            
            // 使用时间戳作为唯一ID
            long uniqueId = System.currentTimeMillis() % 10000;
            
            // 在段落开始处创建书签开始标记
            CTBookmark bookmarkStart = ctp.addNewBookmarkStart();
            bookmarkStart.setName(bookmarkName);
            bookmarkStart.setId(BigInteger.valueOf(uniqueId));
            
            // 在段落结束处创建书签结束标记
            CTMarkupRange bookmarkEnd = ctp.addNewBookmarkEnd();
            bookmarkEnd.setId(BigInteger.valueOf(uniqueId));
            
            System.out.println("✅ 书签 '" + bookmarkName + "' 已创建，ID: " + uniqueId);
            
        } catch (Exception e) {
            System.err.println("创建书签失败: " + e.getMessage());
            // 如果创建书签失败，至少添加文本作为备选
            XWPFRun run = paragraph.createRun();
            run.setText("[" + bookmarkName + "]");
            run.setBold(true);
        }
    }
    
    public static void main(String[] args) {
        try {
            createTestDocument("src/main/resources/doc/test_introduction.docx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
