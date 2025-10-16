package com.example.myjavalab.utils;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.AfterEach;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.*;

public class DocxUtilsTest {

    private String testDir;
    private String originalDocPath;
    private String tempDocPath;
    private String resultDocPath;

    @BeforeEach
    void setUp() {
        testDir = "src/main/resources/doc";
        originalDocPath = "src/main/resources/doc/debug_introduction.docx";
        tempDocPath = testDir + "/temp_introduction.docx";
        resultDocPath = testDir + "/result_introduction.docx";
        
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
    //     try {
    //         // Files.deleteIfExists(Paths.get(tempDocPath));
    //         // 注释掉删除结果文件，保留用于验证
    //         // Files.deleteIfExists(Paths.get(resultDocPath));
    //     } catch (IOException e) {
    //         System.err.println("清理测试文件失败: " + e.getMessage());
    //     }
    // }

    @Test
    void testInsertBookmarkBefore() {
        try {
            // 测试用例1: 在文件中书签labelA之前插入labelB
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            
            // 验证文件是否创建成功
            assertTrue(Files.exists(Paths.get(tempDocPath)), "临时文档应该被创建");
            
            System.out.println("✅ 测试用例1通过: 在labelA之前成功插入labelB");
            
        } catch (Exception e) {
            fail("测试用例1失败: " + e.getMessage());
        }
    }

    @Test
    void testCopyBookmarkContent() {
        try {
            // 先创建临时文档
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            
            // 获取原始文档中labelA的内容
            String originalLabelAContent = DocxUtils.getBookmarkContentFromFile(originalDocPath, "labelA");
            System.out.println("📝 原始文档labelA内容: '" + originalLabelAContent + "'");
            
            // 测试用例2: 将labelA的内容复制给labelB
            DocxUtils.copyBookmarkContent(tempDocPath, resultDocPath, "labelA", "labelB");
            
            // 验证文件是否创建成功
            assertTrue(Files.exists(Paths.get(resultDocPath)), "结果文档应该被创建");
            
            // 验证源文件和目标文件labelA内容一致性
            String tempLabelAContent = DocxUtils.getBookmarkContentFromFile(tempDocPath, "labelA");
            String resultLabelAContent = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelA");
            
            System.out.println("📝 临时文档labelA内容: '" + tempLabelAContent + "'");
            System.out.println("📝 结果文档labelA内容: '" + resultLabelAContent + "'");
            
            // 验证labelA内容在复制前后保持一致
            assertEquals(originalLabelAContent, tempLabelAContent, "临时文档中labelA内容应该与原始文档一致");
            assertEquals(originalLabelAContent, resultLabelAContent, "结果文档中labelA内容应该与原始文档一致");
            
            // 验证result_introduction里labelA和labelB内容一致性
            String resultLabelBContent = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelB");
            System.out.println("📝 结果文档labelB内容: '" + resultLabelBContent + "'");
            
            assertEquals(originalLabelAContent, resultLabelBContent, "结果文档中labelB内容应该与原始labelA内容一致");
            
            System.out.println("✅ 测试用例2通过: 成功将labelA的内容复制给labelB，内容验证通过");
            
        } catch (Exception e) {
            fail("测试用例2失败: " + e.getMessage());
        }
    }

    @Test
    void testCompleteWorkflow() {
        try {
            // 完整工作流程测试
            System.out.println("开始完整工作流程测试...");
            
            // 步骤1: 在labelA之前插入labelB
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            System.out.println("步骤1完成: 在labelA之前插入labelB");
            
            // 步骤2: 将labelA的内容复制给labelB
            DocxUtils.copyBookmarkContent(tempDocPath, resultDocPath, "labelA", "labelB");
            System.out.println("步骤2完成: 将labelA的内容复制给labelB");
            
            // 步骤3: 验证结果
            assertTrue(Files.exists(Paths.get(resultDocPath)), "结果文档应该存在");
            
            // 检查文件大小是否合理
            long fileSize = Files.size(Paths.get(resultDocPath));
            assertTrue(fileSize > 0, "结果文档不应该为空");
            
            System.out.println("✅ 完整工作流程测试通过");
            System.out.println("📁 结果文档路径: " + resultDocPath);
            System.out.println("📊 文件大小: " + fileSize + " bytes");
            
        } catch (Exception e) {
            fail("完整工作流程测试失败: " + e.getMessage());
        }
    }

    @Test
    void testErrorHandling() {
        // 测试错误处理
        assertThrows(IllegalArgumentException.class, () -> {
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "nonExistentBookmark", "labelB");
        }, "应该抛出异常当书签不存在时");
        
        System.out.println("✅ 错误处理测试通过");
    }

    @Test
    void testFileCreationAndVerification() {
        try {
            // 创建临时文档
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            
            // 创建结果文档
            DocxUtils.copyBookmarkContent(tempDocPath, resultDocPath, "labelA", "labelB");
            
            // 验证文件存在
            assertTrue(Files.exists(Paths.get(tempDocPath)), "临时文档应该存在");
            assertTrue(Files.exists(Paths.get(resultDocPath)), "结果文档应该存在");
            
            // 验证文件大小
            long tempSize = Files.size(Paths.get(tempDocPath));
            long resultSize = Files.size(Paths.get(resultDocPath));
            
            assertTrue(tempSize > 0, "临时文档不应该为空");
            assertTrue(resultSize > 0, "结果文档不应该为空");
            
            System.out.println("✅ 文件创建和验证测试通过");
            System.out.println("📁 临时文档: " + tempDocPath + " (大小: " + tempSize + " bytes)");
            System.out.println("📁 结果文档: " + resultDocPath + " (大小: " + resultSize + " bytes)");
            
        } catch (Exception e) {
            fail("文件创建和验证测试失败: " + e.getMessage());
        }
    }
}
