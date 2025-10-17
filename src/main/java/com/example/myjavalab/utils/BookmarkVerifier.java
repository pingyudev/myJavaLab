package com.example.myjavalab.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * 验证最终结果文档中的书签
 */
public class BookmarkVerifier {

    public static void main(String[] args) {
        try {
            System.out.println("=== 书签验证程序 ===");
            
            // 检查各个阶段的文档
            String[] docs = {
                "src/main/resources/doc/debug_introduction.docx",
                "src/main/resources/doc/temp_introduction.docx", 
                "src/main/resources/doc/result_introduction.docx"
            };
            
            String[] names = {"原始文档", "临时文档", "结果文档"};
            
            for (int i = 0; i < docs.length; i++) {
                System.out.println("\n=== " + names[i] + " ===");
                verifyDocument(docs[i]);
            }
            
        } catch (Exception e) {
            System.err.println("验证失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private static void verifyDocument(String docPath) throws IOException {
        try (FileInputStream fis = new FileInputStream(docPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            
            System.out.println("文档路径: " + docPath);
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            System.out.println("段落数量: " + paragraphs.size());
            
            int totalBookmarks = 0;
            for (int i = 0; i < paragraphs.size(); i++) {
                XWPFParagraph paragraph = paragraphs.get(i);
                System.out.println("段落 " + i + ": " + paragraph.getText());
                
                try {
                    CTP ctp = paragraph.getCTP();
                    CTBookmark[] bookmarks = ctp.getBookmarkStartArray();
                    if (bookmarks.length > 0) {
                        System.out.println("  📌 书签数量: " + bookmarks.length);
                        for (CTBookmark bookmark : bookmarks) {
                            System.out.println("  📌 书签名称: " + bookmark.getName());
                            System.out.println("  📌 书签ID: " + bookmark.getId());
                            totalBookmarks++;
                        }
                    }
                } catch (Exception e) {
                    System.out.println("  ❌ 无法访问书签: " + e.getMessage());
                }
            }
            
            System.out.println("📊 总书签数量: " + totalBookmarks);
        }
    }
}
