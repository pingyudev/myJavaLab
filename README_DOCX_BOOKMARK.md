# DOCX书签操作工具

基于Apache POI实现的DOCX文档书签操作功能。

## 功能特性

1. **在指定书签A前面插入新书签B** - `insertBookmarkBefore()`
2. **将书签A的内容复制到书签B** - `copyBookmarkContent()`
3. **多次复制指定书签内容** - `copyBookmarkContentMultipleTimes()`

## 项目结构

```
src/main/java/com/example/myjavalab/utils/
├── DocxUtils.java                    # 核心书签操作工具类
├── DocxTestDocumentCreator.java      # 测试文档创建器
└── DocxBookmarkDemo.java             # 演示程序

src/test/java/com/example/myjavalab/utils/
└── DocxUtilsTest.java                # 单元测试

src/main/resources/doc/
├── introduction.docx                  # 原始文档
├── demo_introduction.docx             # 演示用测试文档
├── temp_introduction.docx             # 临时处理文档
└── result_introduction.docx           # 最终结果文档
```

## 使用方法

### 1. 基本用法

```java
// 在labelA之前插入labelB
DocxUtils.insertBookmarkBefore(
    "input.docx", 
    "output.docx", 
    "labelA", 
    "labelB"
);

// 将labelA的内容复制给labelB
DocxUtils.copyBookmarkContent(
    "input.docx", 
    "output.docx", 
    "labelA", 
    "labelB"
);

// 多次复制labelA的内容（创建labelA1, labelA2, labelA3等）
DocxUtils.copyBookmarkContentMultipleTimes(
    "source.docx", 
    "result.docx", 
    "labelA", 
    3
);
```

### 2. 运行演示

```bash
# 运行演示程序
./mvnw exec:java -Dexec.mainClass="com.example.myjavalab.utils.DocxBookmarkDemo"

# 运行单元测试
./mvnw test -Dtest=DocxUtilsTest
```

### 3. 测试用例

演示程序包含以下测试用例：

1. **测试用例1**: 在文件中书签labelA之前插入labelB
2. **测试用例2**: 将labelA的内容复制给labelB
3. **测试用例3**: 在同目录下生成文档副本，检查副本中是否已经得到想要的结果
4. **测试用例4**: 多次复制指定书签内容（创建labelA1, labelA2, labelA3等）

## 技术实现

### 核心方法

- `findBookmarkPosition()` - 查找书签在文档中的位置
- `insertBookmarkAtPosition()` - 在指定位置插入书签
- `getBookmarkContent()` - 获取书签的内容
- `setBookmarkContent()` - 设置书签的内容
- `copyBookmarkContentMultipleTimes()` - 多次复制指定书签内容

### 依赖库

- Apache POI 5.2.4 (poi-ooxml, poi-scratchpad)
- Spring Boot 2.7.18
- JUnit 5

## 测试结果

✅ 所有8个测试用例通过
✅ 成功生成文档副本
✅ 书签操作功能正常工作
✅ 多次复制功能正常工作

## 注意事项

1. 确保输入的DOCX文档包含指定的书签
2. 书签名称区分大小写
3. 操作会创建新的文档文件，不会修改原始文档
4. 支持JDK 8及以上版本

## 错误处理

- 当书签不存在时，会抛出 `IllegalArgumentException`
- 文件操作异常会抛出 `IOException`
- 文档格式错误会抛出相应的POI异常
