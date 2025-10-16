package com.example.myjavalab.utils;

import java.io.IOException;

/**
 * 多次复制书签内容演示程序
 */
public class MultipleCopyDemo {

    public static void main(String[] args) {
        try {
            System.out.println("=== 多次复制书签内容演示 ===");
            
            // 使用现有的测试文档作为源文件
            String sourceFile = "debug_introduction.docx";
            String targetFile = "result_introduction.docx";
            String sourceLabel = "labelA";
            int copyTimes = 3;
            
            System.out.println("📄 源文件: " + sourceFile);
            System.out.println("📄 目标文件: " + targetFile);
            System.out.println("🏷️ 源书签: " + sourceLabel);
            System.out.println("🔢 复制次数: " + copyTimes);
            System.out.println();
            
            // 执行多次复制操作
            DocxUtils.copyBookmarkContentMultipleTimes(sourceFile, targetFile, sourceLabel, copyTimes);
            
            System.out.println("\n=== 演示完成 ===");
            System.out.println("请检查生成的文件: src/main/resources/doc/" + targetFile);
            
        } catch (Exception e) {
            System.err.println("演示失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
