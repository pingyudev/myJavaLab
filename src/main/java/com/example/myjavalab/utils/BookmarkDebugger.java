package com.example.myjavalab.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

public class BookmarkDebugger {
    
    public static void main(String[] args) {
        try {
            String documentPath = "src/main/resources/doc/debug_introduction.docx";
            debugBookmarks(documentPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    public static void debugBookmarks(String documentPath) throws IOException {
        try (FileInputStream fis = new FileInputStream(documentPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            
            System.out.println("🔍 调试文档中的书签位置...");
            System.out.println("📄 文档段落总数: " + document.getParagraphs().size());
            
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (int i = 0; i < paragraphs.size(); i++) {
                XWPFParagraph paragraph = paragraphs.get(i);
                String text = paragraph.getText();
                System.out.println("\n📝 段落 " + i + ": '" + text + "'");
                
                // 检查书签
                try {
                    CTP ctp = paragraph.getCTP();
                    CTBookmark[] bookmarks = ctp.getBookmarkStartArray();
                    if (bookmarks != null && bookmarks.length > 0) {
                        System.out.println("  📌 包含书签:");
                        for (CTBookmark bookmark : bookmarks) {
                            System.out.println("    - 书签名称: " + bookmark.getName());
                            System.out.println("    - 书签ID: " + bookmark.getId());
                        }
                    } else {
                        System.out.println("  ❌ 无书签");
                    }
                } catch (Exception e) {
                    System.out.println("  ⚠️ 检查书签时出错: " + e.getMessage());
                }
            }
        }
    }
}