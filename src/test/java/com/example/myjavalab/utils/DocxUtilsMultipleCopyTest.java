package com.example.myjavalab.utils;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.AfterEach;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.*;

public class DocxUtilsMultipleCopyTest {

    private String testDir;
    private String originalDocPath;
    private String resultDocPath;

    @BeforeEach
    void setUp() {
        testDir = "src/test/resources/test-output";
        originalDocPath = "src/main/resources/doc/debug_introduction.docx";
        resultDocPath = "src/main/resources/doc/result_introduction.docx";
        
        // 创建测试目录
        try {
            Files.createDirectories(Paths.get(testDir));
        } catch (IOException e) {
            fail("无法创建测试目录: " + e.getMessage());
        }
    }

    @AfterEach
    void tearDown() {
        // 清理测试文件
        try {
            Files.deleteIfExists(Paths.get(resultDocPath));
        } catch (IOException e) {
            System.err.println("清理测试文件失败: " + e.getMessage());
        }
    }

    @Test
    void testCopyBookmarkContentMultipleTimes() {
        try {
            // 测试多次复制书签内容
            DocxUtils.copyBookmarkContentMultipleTimes(
                "debug_introduction.docx", 
                "result_introduction.docx", 
                "labelA", 
                3
            );
            
            // 验证文件是否创建成功
            assertTrue(Files.exists(Paths.get(resultDocPath)), "结果文档应该被创建");
            
            // 检查文件大小是否合理
            long fileSize = Files.size(Paths.get(resultDocPath));
            assertTrue(fileSize > 0, "结果文档不应该为空");
            
            System.out.println("✅ 多次复制书签内容测试通过");
            System.out.println("📁 结果文档路径: " + resultDocPath);
            System.out.println("📊 文件大小: " + fileSize + " bytes");
            
        } catch (Exception e) {
            System.err.println("测试失败详情: " + e.getMessage());
            e.printStackTrace();
            fail("多次复制书签内容测试失败: " + e.getMessage());
        }
    }

    @Test
    void testCopyBookmarkContentMultipleTimesWithInvalidBookmark() {
        // 测试使用不存在的书签
        assertThrows(IllegalArgumentException.class, () -> {
            DocxUtils.copyBookmarkContentMultipleTimes(
                "debug_introduction.docx", 
                "result_introduction.docx", 
                "nonExistentBookmark", 
                2
            );
        }, "应该抛出异常当书签不存在时");
        
        System.out.println("✅ 无效书签错误处理测试通过");
    }

    @Test
    void testCopyBookmarkContentMultipleTimesWithZeroCopies() {
        try {
            // 测试复制次数为0的情况
            DocxUtils.copyBookmarkContentMultipleTimes(
                "debug_introduction.docx", 
                "result_introduction.docx", 
                "labelA", 
                0
            );
            
            // 验证文件是否创建成功
            assertTrue(Files.exists(Paths.get(resultDocPath)), "结果文档应该被创建");
            
            System.out.println("✅ 零次复制测试通过");
            
        } catch (Exception e) {
            fail("零次复制测试失败: " + e.getMessage());
        }
    }
}
