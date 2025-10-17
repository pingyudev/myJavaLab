package com.example.myjavalab.utils;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.AfterEach;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.*;

/**
 * 测试样式保持和序号同步功能
 */
public class StyleAndNumberingTest {

    private String testDir;
    private String originalDocPath;
    private String tempDocPath;
    private String resultDocPath;

    @BeforeEach
    void setUp() {
        testDir = "src/main/resources/doc";
        originalDocPath = "src/main/resources/doc/debug_introduction.docx";
        tempDocPath = testDir + "/temp_style_test.docx";
        resultDocPath = testDir + "/result_style_test.docx";
        
        // 创建测试目录
        try {
            Files.createDirectories(Paths.get(testDir));
            
            // 如果测试文档不存在，创建一个
            if (!Files.exists(Paths.get(originalDocPath))) {
                DocxTestDocumentCreator.createTestDocument(originalDocPath);
            }
        } catch (IOException e) {
            fail("无法创建测试目录或文档: " + e.getMessage());
        }
    }

    @AfterEach
    void tearDown() {
        // 清理测试文件（保留结果文件用于验证）
    }

    @Test
    void testStylePreservationAndNumbering() {
        try {
            System.out.println("开始测试样式保持和序号同步...");
            
            // 步骤1: 在labelA之前插入labelB
            System.out.println("步骤1: 在labelA之前插入labelB");
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            
            // 验证文件是否创建成功
            assertTrue(Files.exists(Paths.get(tempDocPath)), "临时文档应该被创建");
            
            // 步骤2: 将labelA的内容复制给labelB
            System.out.println("步骤2: 将labelA的内容复制给labelB");
            DocxUtils.copyBookmarkContent(tempDocPath, resultDocPath, "labelA", "labelB");
            
            // 验证文件是否创建成功
            assertTrue(Files.exists(Paths.get(resultDocPath)), "结果文档应该被创建");
            
            // 步骤3: 验证序号同步
            System.out.println("步骤3: 验证序号同步");
            verifyNumberingSync();
            
            // 步骤4: 验证样式保持
            System.out.println("步骤4: 验证样式保持");
            verifyStylePreservation();
            
            System.out.println("✅ 样式保持和序号同步测试通过");
            
        } catch (Exception e) {
            fail("样式保持和序号同步测试失败: " + e.getMessage());
        }
    }
    
    /**
     * 验证序号同步
     */
    private void verifyNumberingSync() throws Exception {
        // 检查原始文档中labelA的序号
        String originalContent = DocxUtils.getBookmarkContentFromFile(originalDocPath, "labelA");
        System.out.println("📝 原始文档labelA内容: '" + originalContent + "'");
        
        // 检查结果文档中labelB的序号（应该是2.）
        String labelBContent = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelB");
        System.out.println("📝 结果文档labelB内容: '" + labelBContent + "'");
        
        // 检查结果文档中labelA的序号（应该是3.）
        String labelAContent = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelA");
        System.out.println("📝 结果文档labelA内容: '" + labelAContent + "'");
        
        // 验证labelB内容不包含序号（因为序号是单独添加的）
        assertFalse(labelBContent.startsWith("2. ") || labelBContent.startsWith("3. "), 
            "labelB内容不应该包含序号，实际内容: " + labelBContent);
        
        // 验证labelA使用编号样式（现在使用Word编号样式，不包含文本序号）
        boolean labelAHasNumbering = DocxUtils.isBookmarkUsingNumberingStyle(resultDocPath, "labelA");
        assertTrue(labelAHasNumbering, "labelA应该使用Word编号样式");
    }
    
    /**
     * 验证样式保持
     */
    private void verifyStylePreservation() throws Exception {
        // 检查labelB内容是否包含粗体部分
        String labelBContent = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelB");
        
        // 验证包含粗体标题
        assertTrue(labelBContent.contains("提升职场竞争力，拥抱AI浪潮："), 
            "labelB应该包含粗体标题部分");
        
        // 验证包含详细内容
        assertTrue(labelBContent.contains("当前大型科技公司偏好高学历人才"), 
            "labelB应该包含详细内容");
        
        System.out.println("✅ 样式保持验证通过");
    }

    @Test
    void testMultipleInsertions() {
        try {
            System.out.println("开始测试多次插入...");
            
            // 第一次插入
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            DocxUtils.copyBookmarkContent(tempDocPath, resultDocPath, "labelA", "labelB");
            
            // 第二次插入（在labelA之前插入labelC）
            String tempDocPath2 = testDir + "/temp_style_test2.docx";
            String resultDocPath2 = testDir + "/result_style_test2.docx";
            
            DocxUtils.insertBookmarkBefore(resultDocPath, tempDocPath2, "labelA", "labelC");
            DocxUtils.copyBookmarkContent(tempDocPath2, resultDocPath2, "labelA", "labelC");
            
            // 验证序号
            String labelBContent = DocxUtils.getBookmarkContentFromFile(resultDocPath2, "labelB");
            String labelCContent = DocxUtils.getBookmarkContentFromFile(resultDocPath2, "labelC");
            String labelAContent = DocxUtils.getBookmarkContentFromFile(resultDocPath2, "labelA");
            
            System.out.println("📝 多次插入后labelB内容: '" + labelBContent + "'");
            System.out.println("📝 多次插入后labelC内容: '" + labelCContent + "'");
            System.out.println("📝 多次插入后labelA内容: '" + labelAContent + "'");
            
            // 验证内容不包含序号（因为序号是单独添加的）
            assertFalse(labelBContent.startsWith("2. ") || labelBContent.startsWith("3. ") || labelBContent.startsWith("4. "), 
                "labelB内容不应该包含序号，实际内容: " + labelBContent);
            assertFalse(labelCContent.startsWith("2. ") || labelCContent.startsWith("3. ") || labelCContent.startsWith("4. "), 
                "labelC内容不应该包含序号，实际内容: " + labelCContent);
            // 验证labelA使用编号样式（现在使用Word编号样式，不包含文本序号）
            boolean labelAHasNumbering = DocxUtils.isBookmarkUsingNumberingStyle(resultDocPath2, "labelA");
            assertTrue(labelAHasNumbering, "labelA应该使用Word编号样式");
            
            System.out.println("✅ 多次插入测试通过");
            
        } catch (Exception e) {
            fail("多次插入测试失败: " + e.getMessage());
        }
    }
}
