package com.example.myjavalab.utils;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.AfterEach;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.*;

public class DocxUtilsTest {

    private String testDir;
    private String originalDocPath;
    private String tempDocPath;
    private String resultDocPath;

    @BeforeEach
    void setUp() {
        testDir = "src/test/resources/test-output";
        originalDocPath = "src/main/resources/doc/debug_introduction.docx";
        tempDocPath = testDir + "/temp_introduction.docx";
        resultDocPath = testDir + "/result_introduction.docx";
        
        // 创建测试目录
        try {
            Files.createDirectories(Paths.get(testDir));
            
            // 如果测试文档不存在，创建一个
            if (!Files.exists(Paths.get(originalDocPath))) {
                // DocxTestDocumentCreator.createTestDocument(originalDocPath);
                throw new RuntimeException("测试文档不存在");
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
    void testInsertBookmarkBefore() {
        try {
            // 获取原始文档中labelA的位置和范围
            int originalLabelAPosition = DocxUtils.getBookmarkPositionFromFile(originalDocPath, "labelA");
            BookmarkRange originalLabelARange = DocxUtils.getBookmarkRangeFromFile(originalDocPath, "labelA");
            System.out.println("📝 原始文档中labelA位置: " + originalLabelAPosition);
            System.out.println("📝 原始文档中labelA范围: " + originalLabelARange);
            
            // 测试用例1: 在文件中书签labelA之前插入labelB
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            
            // 验证文件是否创建成功
            assertTrue(Files.exists(Paths.get(tempDocPath)), "临时文档应该被创建");
            
            // 验证插入后的位置顺序
            int newLabelAPosition = DocxUtils.getBookmarkPositionFromFile(tempDocPath, "labelA");
            int newLabelBPosition = DocxUtils.getBookmarkPositionFromFile(tempDocPath, "labelB");
            BookmarkRange newLabelARange = DocxUtils.getBookmarkRangeFromFile(tempDocPath, "labelA");
            BookmarkRange newLabelBRange = DocxUtils.getBookmarkRangeFromFile(tempDocPath, "labelB");
            
            System.out.println("📝 插入后labelA位置: " + newLabelAPosition);
            System.out.println("📝 插入后labelB位置: " + newLabelBPosition);
            System.out.println("📝 插入后labelA范围: " + newLabelARange);
            System.out.println("📝 插入后labelB范围: " + newLabelBRange);
            
            // 验证书签范围有效
            assertTrue(newLabelARange.isValid(), "labelA书签范围应该有效");
            assertTrue(newLabelBRange.isValid(), "labelB书签范围应该有效");
            
            // 验证labelB确实插入到了labelA之前
            assertTrue(newLabelBPosition < newLabelAPosition, 
                "labelB应该插入到labelA之前，但实际位置: labelB=" + newLabelBPosition + ", labelA=" + newLabelAPosition);
            
            // 验证labelA的位置向后移动了一位（因为插入了新段落）
            assertEquals(originalLabelAPosition + 1, newLabelAPosition, 
                "labelA的位置应该向后移动一位");
            
            // 验证labelB的位置就是原来labelA的位置
            assertEquals(originalLabelAPosition, newLabelBPosition, 
                "labelB应该插入到原来labelA的位置");
            
            // 验证labelB的内容包含initialString（说明书签正确包围了内容）
            String labelBContent = DocxUtils.getBookmarkContentFromFile(tempDocPath, "labelB");
            assertTrue(labelBContent.contains("initialString"), 
                "labelB书签应该包围initialString内容，实际内容: " + labelBContent);
            
            System.out.println("📝 labelB书签内容: '" + labelBContent + "'");
            System.out.println("✅ 测试用例1通过: 在labelA之前成功插入labelB，位置验证通过，书签内容验证通过");
            
        } catch (Exception e) {
            fail("测试用例1失败: " + e.getMessage());
        }
    }

    @Test
    void testCopyBookmarkContent() {
        try {
            // 先测试原始文档中的书签内容提取
            String originalLabelAContent = DocxUtils.getBookmarkContentFromFile(originalDocPath, "labelA");
            System.out.println("📝 原始文档labelA内容: '" + originalLabelAContent + "'");
            
            // 如果原始文档中labelA内容为空，直接失败测试
            if (originalLabelAContent == null || originalLabelAContent.trim().isEmpty()) {
                fail("原始文档中labelA内容为空，无法进行复制测试");
            }
            
            // 先创建临时文档
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            
            // 测试用例2: 将labelA的内容复制给labelB
            DocxUtils.copyBookmarkContent(tempDocPath, resultDocPath, "labelA", "labelB");
            
            // 验证文件是否创建成功
            assertTrue(Files.exists(Paths.get(resultDocPath)), "结果文档应该被创建");
            
            // 验证源文件和目标文件labelA内容一致性
            String tempLabelAContent = DocxUtils.getBookmarkContentFromFile(tempDocPath, "labelA");
            String resultLabelAContent = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelA");
            
            System.out.println("📝 临时文档labelA内容: '" + tempLabelAContent + "'");
            System.out.println("📝 结果文档labelA内容: '" + resultLabelAContent + "'");
            
            // 验证labelA内容在复制前后保持一致（除了序号变化）
            // 移除序号进行比较
            String originalContentWithoutNumber = removeNumberFromContent(originalLabelAContent);
            String tempContentWithoutNumber = removeNumberFromContent(tempLabelAContent);
            String resultContentWithoutNumber = removeNumberFromContent(resultLabelAContent);
            
            assertEquals(originalContentWithoutNumber, tempContentWithoutNumber, "临时文档中labelA内容（除序号）应该与原始文档一致");
            assertEquals(originalContentWithoutNumber, resultContentWithoutNumber, "结果文档中labelA内容（除序号）应该与原始文档一致");
            
            // 验证result_introduction里labelA和labelB内容一致性
            String resultLabelBContent = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelB");
            System.out.println("📝 结果文档labelB内容: '" + resultLabelBContent + "'");
            
            assertEquals(originalContentWithoutNumber, resultLabelBContent, "结果文档中labelB内容应该与原始labelA内容（除序号）一致");
            
            // 验证目标文件中labelA的内容和源文件labelA中的内容一致
            String originalLabelAContentInOriginalDoc = DocxUtils.getBookmarkContentFromFile(originalDocPath, "labelA");
            String resultLabelAContentInResultDoc = DocxUtils.getBookmarkContentFromFile(resultDocPath, "labelA");

            assertEquals(
                removeNumberFromContent(originalLabelAContentInOriginalDoc),
                removeNumberFromContent(resultLabelAContentInResultDoc),
                "目标文件中labelA的内容（除序号）应该和源文件labelA中的内容一致"
            );

            assertNotNull(resultLabelAContent, "结果文档中labelA内容不应为空");
            assertFalse(resultLabelAContent.trim().isEmpty(), "结果文档中labelA内容不应为空字符串");

            // 验证目标文件中的labelA内容和目标文件中的labelB内容一致
            assertEquals(
                removeNumberFromContent(resultLabelAContent),
                removeNumberFromContent(resultLabelBContent),
                "结果文档中labelA和labelB的内容（除序号）应该一致"
            );
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
    
    /**
     * 从内容中移除序号（辅助方法）
     */
    private String removeNumberFromContent(String content) {
        if (content != null && content.matches("^\\d+\\..*")) {
            return content.substring(content.indexOf('.') + 1).trim();
        }
        return content;
    }
    
    @Test
    void testNumberingStylePreservation() {
        try {
            System.out.println("开始测试编号样式保持...");
            
            // 步骤1: 在labelA之前插入labelB
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            
            // 步骤2: 将labelA的内容复制给labelB
            DocxUtils.copyBookmarkContent(tempDocPath, resultDocPath, "labelA", "labelB");
            
            // 步骤3: 验证编号样式
            boolean labelBHasNumbering = DocxUtils.isBookmarkUsingNumberingStyle(resultDocPath, "labelB");
            boolean labelAHasNumbering = DocxUtils.isBookmarkUsingNumberingStyle(resultDocPath, "labelA");
            
            System.out.println("📝 labelB是否使用编号样式: " + labelBHasNumbering);
            System.out.println("📝 labelA是否使用编号样式: " + labelAHasNumbering);
            
            // 验证labelB使用编号样式
            assertTrue(labelBHasNumbering, "labelB应该使用Word编号样式");
            
            // 验证labelA使用编号样式
            assertTrue(labelAHasNumbering, "labelA应该使用Word编号样式");
            
            System.out.println("✅ 编号样式保持测试通过");
            
        } catch (Exception e) {
            fail("编号样式保持测试失败: " + e.getMessage());
        }
    }
    
    @Test
    void testNumberingStyleAfterInsertion() {
        try {
            System.out.println("开始测试插入后的编号样式...");
            
            // 在labelA之前插入labelB
            DocxUtils.insertBookmarkBefore(originalDocPath, tempDocPath, "labelA", "labelB");
            
            // 验证插入后labelB是否使用编号样式
            boolean labelBHasNumbering = DocxUtils.isBookmarkUsingNumberingStyle(tempDocPath, "labelB");
            boolean labelAHasNumbering = DocxUtils.isBookmarkUsingNumberingStyle(tempDocPath, "labelA");
            
            System.out.println("📝 插入后labelB是否使用编号样式: " + labelBHasNumbering);
            System.out.println("📝 插入后labelA是否使用编号样式: " + labelAHasNumbering);
            
            // 验证labelB使用编号样式
            assertTrue(labelBHasNumbering, "插入后labelB应该使用Word编号样式");
            
            // 验证labelA使用编号样式
            assertTrue(labelAHasNumbering, "插入后labelA应该使用Word编号样式");
            
            System.out.println("✅ 插入后编号样式测试通过");
            
        } catch (Exception e) {
            fail("插入后编号样式测试失败: " + e.getMessage());
        }
    }
}
