# My Java Lab - Claude Code Configuration

## 项目概述
Spring Boot 2.7.18 Java项目，专注于DOCX文档处理功能，使用Apache POI库操作Office文档。

## 技术栈
- **框架**: Spring Boot 2.7.18
- **Java版本**: 1.8
- **构建工具**: Maven
- **文档处理**: Apache POI 5.2.4
- **测试**: Spring Boot Test

## 项目结构
```
src/
├── main/java/com/example/myjavalab/
│   ├── MyJavaLabApplication.java    # 应用程序入口
│   └── utils/
│       ├── DocxUtils.java           # DOCX处理工具类
│       └── BookmarkRange.java        # 书签范围处理
└── test/java/com/example/myjavalab/
    └── utils/
        └── DocxUtilsTest.java       # 测试类
```

## 构建和运行命令
```bash
# 编译项目
mvn clean compile

# 运行测试
mvn test

# 打包应用
mvn clean package

# 运行应用程序
mvn spring-boot:run

# 直接运行JAR包
java -jar target/my-java-lab-0.0.1-SNAPSHOT.jar
```

## 代码约定
### Java代码规范
- 使用Spring Boot标准配置
- 遵循Java 8语法特性
- 类名使用PascalCase（如：DocxUtils）
- 方法名使用camelCase（如：processDocument）
- 常量使用UPPER_SNAKE_CASE（如：MAX_LENGTH）

### 依赖管理
- 通过pom.xml管理所有依赖
- 使用Spring Boot Starter简化依赖配置
- POI版本固定为5.2.4以确保兼容性

### 测试规范
- 测试类位于src/test/java
- 使用JUnit和Spring Boot Test
- 测试方法以test开头或@Test注解

### 测试辅助方法管理规则
**原则**: 测试辅助方法应该首先在test包下开发和验证，成熟后再迁移到主代码中

**具体规则**:
1. **辅助方法创建**: 在单元测试过程中涉及的辅助方法，统一放在test包下面
2. **Helper类命名**: 在test包下创建`{原类名}+Helper`的类，例如`DocxUtilsHelper`
3. **方法存放**: 将测试辅助方法放在对应的Helper类中
4. **方法迁移**: 当原类里面的核心方法需要使用Helper类中的方法时，将这些方法迁移到原类里（转正）
5. **迁移时机**: 只有在主代码确实需要复用这些辅助方法时才进行迁移
6. **清理工作**: 方法迁移后，可以从Helper类中删除对应的辅助方法

**示例结构**:
```
src/test/java/com/example/myjavalab/
├── utils/
│   ├── DocxUtilsTest.java       # 测试类
│   └── DocxUtilsHelper.java     # 测试辅助方法类
```

**优势**:
- 保持主代码的简洁性
- 测试辅助方法经过充分验证后再进入生产代码
- 避免过早优化和过度设计
- 提供清晰的测试到生产的方法迁移路径

## 开发指南
### 添加新功能
1. 在src/main/java/com/example/myjavalab/下创建相应包结构
2. 实现业务逻辑类
3. 在对应包下创建测试类
4. 编写单元测试
5. 运行mvn test确保测试通过

### DOCX处理
- 使用DocxUtils类进行文档操作
- 支持书签提取和内容处理
- 处理时注意异常处理和资源释放

### 调试和问题排查
1. 使用System.out.println或日志框架输出调试信息
2. 检查Maven依赖是否正确加载
3. 验证POI库的兼容性
4. 确保文件路径和权限正确

## 常见问题
### Maven相关问题
- 清理缓存：`mvn clean`
- 重新下载依赖：`mvn clean install -U`
- 查看依赖树：`mvn dependency:tree`

### Spring Boot相关问题
- 检查application.properties配置
- 验证Spring Boot版本兼容性
- 查看启动日志排查问题

## 最佳实践
1. **依赖管理**: 统一通过pom.xml管理，避免硬编码版本
2. **异常处理**: 使用try-catch-finally处理异常，确保资源释放
3. **代码复用**: 工具类保持单一职责，提高复用性
4. **测试覆盖**: 为核心功能编写单元测试
5. **文档注释**: 为公共API添加JavaDoc注释
6. **版本控制**: 使用语义化版本号，遵循MAJOR.MINOR.PATCH格式

## 算法注释规范
### 核心方法算法注释要求
**适用范围**: 核心类中的核心方法，特别是包含复杂算法逻辑的方法

**注释模板**:
```java
/**
 * 方法功能简述
 *
 * 算法规则详细说明:
 * 1. 步骤1: 描述算法的第一步逻辑和实现原理
 * 2. 步骤2: 描述算法的第二步逻辑和关键判断
 * 3. 步骤3: 描述算法的第三步和边界条件处理
 *
 * 关键参数说明:
 * - paramName: 参数的作用和取值范围
 *
 * 返回值逻辑:
 * - 返回值的含义和生成规则
 *
 * 异常处理策略:
 * - 可能抛出的异常类型和触发条件
 *
 * 性能考虑:
 * - 时间复杂度分析
 * - 空间复杂度分析
 *
 * 示例场景:
 * - 典型使用案例
 * - 特殊情况处理
 *
 * @param param1 参数1说明
 * @param param2 参数2说明
 * @return 返回值说明
 * @throws IOException 异常说明
 */
```

**需要添加算法注释的核心方法**:
1. `extractParagraphContentBetweenBookmarks` - 多段落内容提取算法
2. `findBookmarkRange` - 书签范围定位算法
3. `copyParagraphStyle` - 样式复制算法
4. `insertMultiParagraphBookmarkBefore` - 多段落书签插入算法
5. `setBookmarkContentFromParagraphContent` - 内容设置算法
6. `compareBookmarkParagraphStyles` - 样式比较算法

**代码更新同步要求**:
- 每当通过模型更新代码后，必须同步更新对应方法的算法注释
- 注释更新检查点: 代码提交前、代码审查时
- 确保注释与实际实现保持一致

## 更新日志
- **当前版本**: 0.0.1-SNAPSHOT
- **主要功能**: DOCX文档处理，书签操作
- **Java版本**: 1.8
- **Spring Boot**: 2.7.18