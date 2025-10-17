package com.example.myjavalab.utils;

import java.io.IOException;

/**
 * 演示样式保持和序号同步功能
 */
public class StyleAndNumberingDemo {
    
    public static void main(String[] args) {
        try {
            System.out.println("🚀 开始演示样式保持和序号同步功能...");
            
            // 确保测试文档存在
            String originalDocPath = "src/main/resources/doc/debug_introduction.docx";
            System.out.println("📝 创建测试文档...");
            DocxTestDocumentCreator.createTestDocument(originalDocPath);
            
            // 步骤1: 在labelA之前插入labelB
            System.out.println("\n📋 步骤1: 在labelA之前插入labelB");
            String tempDocPath = "src/main/resources/doc/temp_style_demo.docx";
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            
            // 步骤2: 将labelA的内容复制给labelB
            System.out.println("📋 步骤2: 将labelA的内容复制给labelB");
            String resultDocPath = "src/main/resources/doc/result_style_demo.docx";
            DocxUtils.copyBookmarkContent(tempDocPath, resultDocPath, "labelA", "labelB");
            
            // 步骤3: 验证结果
            System.out.println("\n🔍 步骤3: 验证结果");
            verifyResults(originalDocPath, resultDocPath);
            
            System.out.println("\n✅ 演示完成！请检查生成的文件:");
            System.out.println("📁 原始文档: " + originalDocPath);
            System.out.println("📁 结果文档: " + resultDocPath);
            
        } catch (Exception e) {
            System.err.println("❌ 演示失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private static void verifyResults(String originalDocPath, String resultDocPath) throws Exception {
        // 检查原始文档内容
        String originalLabelAContent = DocxUtils.getBookmarkContentFromFile(originalDocPath, "labelA");
        System.out.println("📝 原始文档labelA内容: '" + originalLabelAContent + "'");
        
        // 检查结果文档内容
        String resultLabelBContent = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelB");
        String resultLabelAContent = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelA");
        
        System.out.println("📝 结果文档labelB内容: '" + resultLabelBContent + "'");
        System.out.println("📝 结果文档labelA内容: '" + resultLabelAContent + "'");
        
        // 验证序号
        if (!resultLabelBContent.startsWith("2. ") && !resultLabelBContent.startsWith("3. ")) {
            System.out.println("✅ labelB内容正确: 不包含序号");
        } else {
            System.out.println("❌ labelB内容错误: " + resultLabelBContent.substring(0, Math.min(10, resultLabelBContent.length())));
        }
        
        if (resultLabelAContent.startsWith("3. ")) {
            System.out.println("✅ labelA序号正确: 3.");
        } else {
            System.out.println("❌ labelA序号错误: " + resultLabelAContent.substring(0, Math.min(10, resultLabelAContent.length())));
        }
        
        // 验证样式
        if (resultLabelBContent.contains("提升职场竞争力，拥抱AI浪潮：")) {
            System.out.println("✅ labelB包含粗体标题部分");
        } else {
            System.out.println("❌ labelB缺少粗体标题部分");
        }
        
        if (resultLabelAContent.contains("提升职场竞争力，拥抱AI浪潮：")) {
            System.out.println("✅ labelA包含粗体标题部分");
        } else {
            System.out.println("❌ labelA缺少粗体标题部分");
        }
    }
}
