package com.example.myjavalab.utils;

import java.io.IOException;

/**
 * DOCX书签操作演示程序
 */
public class DocxBookmarkDemo {

    public static void main(String[] args) {
        try {
            System.out.println("=== DOCX书签操作演示 ===");
            
            // 创建测试文档
            String originalDoc = "src/main/resources/doc/demo_introduction.docx";
            DocxTestDocumentCreator.createTestDocument(originalDoc);
            System.out.println("✅ 创建测试文档: " + originalDoc);
            
            // 测试用例1: 在labelA之前插入labelB
            String tempDoc = "src/main/resources/doc/temp_introduction.docx";
            DocxUtils.insertBookmarkBefore(originalDoc, tempDoc, "labelA", "labelB");
            System.out.println("✅ 测试用例1完成: 在labelA之前插入labelB");
            
            // 测试用例2: 将labelA的内容复制给labelB
            String resultDoc = "src/main/resources/doc/result_introduction.docx";
            DocxUtils.copyBookmarkContent(tempDoc, resultDoc, "labelA", "labelB");
            System.out.println("✅ 测试用例2完成: 将labelA的内容复制给labelB");
            
            // 验证结果
            java.io.File resultFile = new java.io.File(resultDoc);
            if (resultFile.exists()) {
                System.out.println("✅ 结果文档已生成: " + resultDoc);
                System.out.println("📊 文件大小: " + resultFile.length() + " bytes");
            } else {
                System.out.println("❌ 结果文档未生成");
            }
            
            System.out.println("\n=== 演示完成 ===");
            System.out.println("请检查以下文件:");
            System.out.println("📄 原始文档: " + originalDoc);
            System.out.println("📄 临时文档: " + tempDoc);
            System.out.println("📄 结果文档: " + resultDoc);
            
        } catch (Exception e) {
            System.err.println("演示失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
