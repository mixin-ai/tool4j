# tool4j
常用的一些java工具

## DocxTemplateProcessor.java 
用于替换 docx 中的占位符${...}；docx转pdf 方便打印。
- maven 引入
```xml
        <dependency>
            <groupId>com.aspose</groupId>
            <artifactId>aspose-words</artifactId>
            <version>15.8.0</version>
        </dependency>
        <dependency>
            <groupId>cn.hutool</groupId>
            <artifactId>hutool-all</artifactId>
        </dependency>  <!-- 正则表达式用到，如果不想用可以重新写正则表达式 -->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>5.2.3</version>
        </dependency>

        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>5.2.3</version>
            <exclusions>
                <exclusion>
                    <groupId>commons-io</groupId>
                    <artifactId>commons-io</artifactId>
                </exclusion>
            </exclusions>
        </dependency>
        <dependency>
            <groupId>commons-io</groupId>
            <artifactId>commons-io</artifactId>
            <version>2.11.0</version> <!-- POI 5.2.3 依赖此版本，建议使用匹配版本 -->
        </dependency>

```
