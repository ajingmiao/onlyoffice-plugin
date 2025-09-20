# OnlyOffice Chart Binding Reader (Java POI)

这是一个Java POI工具，用于读取OnlyOffice插件生成的图表绑定数据。它与OnlyOffice插件的JavaScript指纹生成算法完全一致，确保能够正确读取图表绑定信息。

## 功能特性

- ✅ **指纹算法一致性**: 与OnlyOffice插件的JavaScript指纹生成算法完全一致
- ✅ **完整数据读取**: 读取图表的所有绑定数据，包括tag、rid、metadata等
- ✅ **多种查找方式**: 支持按chartId、rid等多种方式查找图表绑定
- ✅ **图表类型检测**: 自动检测图表类型（条形图、折线图、饼图等）
- ✅ **错误处理**: 完善的错误处理和日志输出

## 依赖要求

- Java 8+
- Apache POI 5.2.4+
- Jackson 2.15.2+
- Maven 3.6+

## 快速开始

### 1. 编译项目

```bash
# 进入java目录
cd java

# 编译项目
mvn clean compile

# 打包为可执行JAR
mvn clean package
```

### 2. 基本使用

```java
import com.onlyoffice.chart.ChartBindingReader;
import java.util.List;

// 创建读取器
ChartBindingReader reader = new ChartBindingReader();

// 读取Word文档中的所有图表绑定
List<ChartBindingReader.ChartInfo> chartInfos = reader.readChartBindings("document.docx");

// 处理结果
for (ChartBindingReader.ChartInfo chartInfo : chartInfos) {
    System.out.println("图表索引: " + chartInfo.getIndex());
    System.out.println("图表类型: " + chartInfo.getChartType());
    System.out.println("指纹: " + chartInfo.getFingerprint());

    if (chartInfo.getBinding() != null) {
        ChartBindingReader.ChartBinding binding = chartInfo.getBinding();
        System.out.println("Chart ID: " + binding.getChartId());
        System.out.println("RID: " + binding.getRid());
        System.out.println("Tag: " + binding.getTag());
    }
}
```

### 3. 命令行使用

```bash
# 使用编译后的JAR
java -jar target/chart-binding-reader-1.0.0.jar document.docx

# 或者直接运行
mvn exec:java -Dexec.mainClass="com.onlyoffice.chart.ChartBindingReader" -Dexec.args="document.docx"
```

## API文档

### ChartBindingReader

主要的读取器类，提供以下方法：

#### `readChartBindings(String filePath)`
读取Word文档中的所有图表绑定数据。

**参数:**
- `filePath`: Word文档路径

**返回:**
- `List<ChartInfo>`: 图表信息列表

#### `findBindingByChartId(String filePath, String chartId)`
根据chartId查找特定的图表绑定。

**参数:**
- `filePath`: Word文档路径
- `chartId`: 图表ID

**返回:**
- `ChartBinding`: 图表绑定数据，如果未找到则返回null

#### `getAllBindings(String filePath)`
获取所有图表绑定数据。

**参数:**
- `filePath`: Word文档路径

**返回:**
- `List<ChartBinding>`: 图表绑定数据列表

### 数据结构

#### ChartInfo
图表信息对象，包含：
- `index`: 图表在文档中的索引
- `chartType`: 图表类型
- `fingerprint`: 图表指纹
- `binding`: 绑定数据（如果存在）

#### ChartBinding
图表绑定数据对象，包含：
- `chartId`: 图表ID
- `chartType`: 图表类型
- `fingerprint`: 图表指纹
- `rid`: 业务ID
- `tag`: 标签数据
- `tagData`: 解析后的标签数据
- `boundData`: 绑定的业务数据
- `metadata`: 元数据
- `createdAt`: 创建时间
- `lastUpdated`: 最后更新时间

## 指纹生成算法

与OnlyOffice插件的JavaScript版本保持完全一致：

1. **索引部分** (必须): `idx:N` - 图表在文档中的精确位置
2. **类型部分** (可选): `type:chart-type` - 图表类型
3. **尺寸部分** (可选): `size:WxH` - 图表尺寸

示例指纹: `idx:0|type:bar-chart|size:400x300`

## 支持的图表类型

- `bar-chart`: 条形图/柱状图
- `line-chart`: 折线图
- `pie-chart`: 饼图
- `area-chart`: 面积图
- `scatter-chart`: 散点图
- `radar-chart`: 雷达图
- `chart-generic`: 通用图表类型

## 使用场景

1. **数据分析**: 读取OnlyOffice文档中的图表绑定数据进行分析
2. **业务集成**: 将图表数据集成到业务系统中
3. **数据迁移**: 批量读取和处理图表绑定数据
4. **报表系统**: 自动化读取和处理报表中的图表数据

## 错误处理

所有方法都会抛出相应的异常：
- `IOException`: 文件读取错误
- `InvalidFormatException`: 文档格式错误
- `RuntimeException`: 指纹生成或数据解析错误

建议在使用时进行适当的异常处理。

## 注意事项

1. **文档格式**: 只支持.docx格式的Word文档
2. **指纹一致性**: 必须确保OnlyOffice插件和Java POI使用相同的指纹生成算法
3. **自定义属性**: 图表绑定数据存储在Word文档的自定义属性中
4. **内存使用**: 对于大型文档，建议分批处理以避免内存溢出

## 示例输出

```
=== Java POI图表绑定读取开始 ===
找到 3 个图表绑定属性
📊 发现图表 1: bar-chart
📊 发现图表 2: line-chart
📊 发现图表 3: pie-chart
文档中共有 3 个图表
✅ 添加索引部分: idx:0
✅ 添加图表类型: bar-chart
✅ 添加尺寸信息: size:400x300
🔖 最终生成指纹: idx:0|type:bar-chart|size:400x300
✅ 图表 0 找到绑定数据: chart-idx_0_id_123_type_bar-chart_1a2b3c
=== Java POI图表绑定读取完成 ===

=== 图表绑定读取结果 ===
ChartInfo{index=0, chartType='bar-chart', fingerprint='idx:0|type:bar-chart|size:400x300', hasBinding=true}
  - 绑定数据: ChartBinding{chartId='chart-idx_0_id_123_type_bar-chart_1a2b3c', chartType='bar-chart', rid='sales-2024-q1', fingerprint='idx:0|type:bar-chart|size:400x300'}
  - Tag: {"trackId":"sales-2024-q1","groupFields":["region","product"]}
  - RID: sales-2024-q1

总图表数: 3
有绑定数据的图表数: 3
```