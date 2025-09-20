package com.onlyoffice.chart;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.opc.PackageProperties;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartType;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScatterChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTRadarChart;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.JsonNode;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * Java POI图表绑定数据读取器
 * 与OnlyOffice插件的JavaScript指纹生成算法保持一致
 */
public class ChartBindingReader {

    private static final String BINDING_PREFIX = "chart-binding:";
    private ObjectMapper objectMapper = new ObjectMapper();

    /**
     * 图表绑定信息
     */
    public static class ChartBinding {
        private String chartId;
        private String fingerprint;
        private String chartType;
        private String createdAt;
        private String lastUpdated;
        private Map<String, Object> boundData;
        private Map<String, Object> metadata;
        private String tag;
        private Map<String, Object> tagData;
        private String rid;

        // Getters and Setters
        public String getChartId() { return chartId; }
        public void setChartId(String chartId) { this.chartId = chartId; }

        public String getFingerprint() { return fingerprint; }
        public void setFingerprint(String fingerprint) { this.fingerprint = fingerprint; }

        public String getChartType() { return chartType; }
        public void setChartType(String chartType) { this.chartType = chartType; }

        public String getCreatedAt() { return createdAt; }
        public void setCreatedAt(String createdAt) { this.createdAt = createdAt; }

        public String getLastUpdated() { return lastUpdated; }
        public void setLastUpdated(String lastUpdated) { this.lastUpdated = lastUpdated; }

        public Map<String, Object> getBoundData() { return boundData; }
        public void setBoundData(Map<String, Object> boundData) { this.boundData = boundData; }

        public Map<String, Object> getMetadata() { return metadata; }
        public void setMetadata(Map<String, Object> metadata) { this.metadata = metadata; }

        public String getTag() { return tag; }
        public void setTag(String tag) { this.tag = tag; }

        public Map<String, Object> getTagData() { return tagData; }
        public void setTagData(Map<String, Object> tagData) { this.tagData = tagData; }

        public String getRid() { return rid; }
        public void setRid(String rid) { this.rid = rid; }

        @Override
        public String toString() {
            return String.format("ChartBinding{chartId='%s', chartType='%s', rid='%s', fingerprint='%s'}",
                    chartId, chartType, rid, fingerprint);
        }
    }

    /**
     * 图表信息
     */
    public static class ChartInfo {
        private int index;
        private String chartType;
        private String chartId;
        private int width;
        private int height;
        private String fingerprint;
        private ChartBinding binding;

        // Getters and Setters
        public int getIndex() { return index; }
        public void setIndex(int index) { this.index = index; }

        public String getChartType() { return chartType; }
        public void setChartType(String chartType) { this.chartType = chartType; }

        public String getChartId() { return chartId; }
        public void setChartId(String chartId) { this.chartId = chartId; }

        public int getWidth() { return width; }
        public void setWidth(int width) { this.width = width; }

        public int getHeight() { return height; }
        public void setHeight(int height) { this.height = height; }

        public String getFingerprint() { return fingerprint; }
        public void setFingerprint(String fingerprint) { this.fingerprint = fingerprint; }

        public ChartBinding getBinding() { return binding; }
        public void setBinding(ChartBinding binding) { this.binding = binding; }

        @Override
        public String toString() {
            return String.format("ChartInfo{index=%d, chartType='%s', fingerprint='%s', hasBinding=%s}",
                    index, chartType, fingerprint, binding != null);
        }
    }

    /**
     * 读取Word文档中的所有图表绑定数据
     */
    public List<ChartInfo> readChartBindings(String filePath) throws IOException, InvalidFormatException {
        try (FileInputStream fis = new FileInputStream(filePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            System.out.println("=== Java POI图表绑定读取开始 ===");

            // 1. 读取所有自定义属性
            Map<String, ChartBinding> bindingMap = readCustomProperties(document);
            System.out.println("找到 " + bindingMap.size() + " 个图表绑定属性");

            // 2. 扫描文档中的所有图表
            List<ChartInfo> chartInfos = scanDocumentCharts(document);
            System.out.println("文档中共有 " + chartInfos.size() + " 个图表");

            // 3. 为每个图表生成指纹并匹配绑定数据
            for (ChartInfo chartInfo : chartInfos) {
                String fingerprint = generateChartFingerprint(chartInfo, chartInfos.size());
                chartInfo.setFingerprint(fingerprint);

                // 查找对应的绑定数据
                ChartBinding binding = bindingMap.get(fingerprint);
                if (binding != null) {
                    chartInfo.setBinding(binding);
                    System.out.println("✅ 图表 " + chartInfo.getIndex() + " 找到绑定数据: " + binding.getChartId());
                } else {
                    System.out.println("⚠️ 图表 " + chartInfo.getIndex() + " 未找到绑定数据");
                }
            }

            System.out.println("=== Java POI图表绑定读取完成 ===");
            return chartInfos;
        }
    }

    /**
     * 读取文档自定义属性中的图表绑定数据
     */
    private Map<String, ChartBinding> readCustomProperties(XWPFDocument document) {
        Map<String, ChartBinding> bindingMap = new HashMap<>();

        try {
            POIXMLProperties properties = document.getProperties();
            POIXMLProperties.CustomProperties customProps = properties.getCustomProperties();

            if (customProps != null) {
                org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperties ctProps =
                    customProps.getUnderlyingProperties();

                if (ctProps != null && ctProps.getPropertyList() != null) {
                    for (org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty prop : ctProps.getPropertyList()) {
                        String name = prop.getName();
                        if (name != null && name.startsWith(BINDING_PREFIX)) {
                            String fingerprint = name.substring(BINDING_PREFIX.length());
                            String value = null;

                            // 获取属性值
                            if (prop.getLpwstr() != null) {
                                value = prop.getLpwstr();
                            } else if (prop.getLpstr() != null) {
                                value = prop.getLpstr();
                            }

                            if (value != null && !value.trim().isEmpty()) {
                                try {
                                    // 解析JSON数据
                                    JsonNode jsonNode = objectMapper.readTree(value);
                                    ChartBinding binding = parseChartBinding(jsonNode);
                                    binding.setFingerprint(fingerprint);
                                    bindingMap.put(fingerprint, binding);

                                    System.out.println("📋 读取绑定: " + fingerprint + " -> " + binding.getChartId());
                                } catch (Exception e) {
                                    System.err.println("❌ 解析绑定数据失败: " + fingerprint + " - " + e.getMessage());
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("❌ 读取自定义属性失败: " + e.getMessage());
        }

        return bindingMap;
    }

    /**
     * 解析图表绑定JSON数据
     */
    private ChartBinding parseChartBinding(JsonNode jsonNode) {
        ChartBinding binding = new ChartBinding();

        if (jsonNode.has("chartId")) {
            binding.setChartId(jsonNode.get("chartId").asText());
        }
        if (jsonNode.has("chartType")) {
            binding.setChartType(jsonNode.get("chartType").asText());
        }
        if (jsonNode.has("createdAt")) {
            binding.setCreatedAt(jsonNode.get("createdAt").asText());
        }
        if (jsonNode.has("lastUpdated")) {
            binding.setLastUpdated(jsonNode.get("lastUpdated").asText());
        }
        if (jsonNode.has("tag")) {
            binding.setTag(jsonNode.get("tag").asText());
        }
        if (jsonNode.has("rid")) {
            binding.setRid(jsonNode.get("rid").asText());
        }

        // 解析boundData
        if (jsonNode.has("boundData")) {
            binding.setBoundData(objectMapper.convertValue(jsonNode.get("boundData"), Map.class));
        }

        // 解析metadata
        if (jsonNode.has("metadata")) {
            binding.setMetadata(objectMapper.convertValue(jsonNode.get("metadata"), Map.class));
        }

        // 解析tagData
        if (jsonNode.has("tagData")) {
            binding.setTagData(objectMapper.convertValue(jsonNode.get("tagData"), Map.class));
        }

        return binding;
    }

    /**
     * 扫描文档中的所有图表
     */
    private List<ChartInfo> scanDocumentCharts(XWPFDocument document) {
        List<ChartInfo> chartInfos = new ArrayList<>();
        int chartIndex = 0;

        // 扫描段落中的图表
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                List<XWPFChart> charts = run.getEmbeddedCharts();
                for (XWPFChart chart : charts) {
                    ChartInfo chartInfo = createChartInfo(chart, chartIndex++);
                    chartInfos.add(chartInfo);
                    System.out.println("📊 发现图表 " + chartIndex + ": " + chartInfo.getChartType());
                }
            }
        }

        // 扫描表格中的图表
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        for (XWPFRun run : paragraph.getRuns()) {
                            List<XWPFChart> charts = run.getEmbeddedCharts();
                            for (XWPFChart chart : charts) {
                                ChartInfo chartInfo = createChartInfo(chart, chartIndex++);
                                chartInfos.add(chartInfo);
                                System.out.println("📊 发现表格图表 " + chartIndex + ": " + chartInfo.getChartType());
                            }
                        }
                    }
                }
            }
        }

        return chartInfos;
    }

    /**
     * 创建图表信息对象
     */
    private ChartInfo createChartInfo(XWPFChart chart, int index) {
        ChartInfo chartInfo = new ChartInfo();
        chartInfo.setIndex(index);

        // 检测图表类型
        String chartType = detectChartType(chart);
        chartInfo.setChartType(chartType);

        // 获取图表尺寸（如果可用）
        try {
            // POI中获取图表尺寸比较复杂，这里使用默认值
            chartInfo.setWidth(400);
            chartInfo.setHeight(300);
        } catch (Exception e) {
            chartInfo.setWidth(400);
            chartInfo.setHeight(300);
        }

        return chartInfo;
    }

    /**
     * 检测图表类型（与JavaScript版本保持一致）
     */
    private String detectChartType(XWPFChart chart) {
        try {
            CTChart ctChart = chart.getCTChart();
            if (ctChart != null && ctChart.getPlotArea() != null) {

                // 检测条形图
                if (ctChart.getPlotArea().getBarChartList() != null &&
                    !ctChart.getPlotArea().getBarChartList().isEmpty()) {
                    return "bar-chart";
                }

                // 检测折线图
                if (ctChart.getPlotArea().getLineChartList() != null &&
                    !ctChart.getPlotArea().getLineChartList().isEmpty()) {
                    return "line-chart";
                }

                // 检测饼图
                if (ctChart.getPlotArea().getPieChartList() != null &&
                    !ctChart.getPlotArea().getPieChartList().isEmpty()) {
                    return "pie-chart";
                }

                // 检测面积图
                if (ctChart.getPlotArea().getAreaChartList() != null &&
                    !ctChart.getPlotArea().getAreaChartList().isEmpty()) {
                    return "area-chart";
                }

                // 检测散点图
                if (ctChart.getPlotArea().getScatterChartList() != null &&
                    !ctChart.getPlotArea().getScatterChartList().isEmpty()) {
                    return "scatter-chart";
                }

                // 检测雷达图
                if (ctChart.getPlotArea().getRadarChartList() != null &&
                    !ctChart.getPlotArea().getRadarChartList().isEmpty()) {
                    return "radar-chart";
                }
            }
        } catch (Exception e) {
            System.err.println("图表类型检测失败: " + e.getMessage());
        }

        return "chart-generic";
    }

    /**
     * 生成图表指纹（与JavaScript版本完全一致）
     */
    private String generateChartFingerprint(ChartInfo chartInfo, int totalCharts) {
        List<String> parts = new ArrayList<>();

        // 1) 精确位置索引（必须，与JavaScript版本一致）
        parts.add("idx:" + chartInfo.getIndex());
        System.out.println("✅ 添加索引部分: idx:" + chartInfo.getIndex());

        // 2) 图表类型（可选，增强唯一性）
        if (chartInfo.getChartType() != null && !chartInfo.getChartType().equals("error")) {
            parts.add("type:" + chartInfo.getChartType());
            System.out.println("✅ 添加图表类型: " + chartInfo.getChartType());
        }

        // 3) 尺寸信息（可选，与JavaScript版本保持一致）
        if (chartInfo.getWidth() > 0 && chartInfo.getHeight() > 0) {
            String sizePart = "size:" + chartInfo.getWidth() + "x" + chartInfo.getHeight();
            parts.add(sizePart);
            System.out.println("✅ 添加尺寸信息: " + sizePart);
        }

        // 必须至少有索引（严格模式）
        if (parts.isEmpty()) {
            throw new RuntimeException("无法生成任何有效的指纹部分");
        }

        String fingerprint = String.join("|", parts);
        System.out.println("🔖 最终生成指纹: " + fingerprint);
        return fingerprint;
    }

    /**
     * 根据chartId查找图表绑定数据
     */
    public ChartBinding findBindingByChartId(String filePath, String chartId) throws IOException, InvalidFormatException {
        List<ChartInfo> chartInfos = readChartBindings(filePath);

        for (ChartInfo chartInfo : chartInfos) {
            if (chartInfo.getBinding() != null &&
                chartId.equals(chartInfo.getBinding().getChartId())) {
                return chartInfo.getBinding();
            }
        }

        return null;
    }

    /**
     * 获取所有图表绑定数据
     */
    public List<ChartBinding> getAllBindings(String filePath) throws IOException, InvalidFormatException {
        List<ChartInfo> chartInfos = readChartBindings(filePath);
        List<ChartBinding> bindings = new ArrayList<>();

        for (ChartInfo chartInfo : chartInfos) {
            if (chartInfo.getBinding() != null) {
                bindings.add(chartInfo.getBinding());
            }
        }

        return bindings;
    }

    /**
     * 主方法 - 测试用例
     */
    public static void main(String[] args) {
        if (args.length != 1) {
            System.out.println("用法: java ChartBindingReader <word文档路径>");
            return;
        }

        String filePath = args[0];
        ChartBindingReader reader = new ChartBindingReader();

        try {
            System.out.println("开始读取Word文档: " + filePath);

            // 读取所有图表绑定
            List<ChartInfo> chartInfos = reader.readChartBindings(filePath);

            System.out.println("\n=== 图表绑定读取结果 ===");
            for (ChartInfo chartInfo : chartInfos) {
                System.out.println(chartInfo);
                if (chartInfo.getBinding() != null) {
                    ChartBinding binding = chartInfo.getBinding();
                    System.out.println("  - 绑定数据: " + binding);
                    if (binding.getTag() != null) {
                        System.out.println("  - Tag: " + binding.getTag());
                    }
                    if (binding.getRid() != null) {
                        System.out.println("  - RID: " + binding.getRid());
                    }
                }
                System.out.println();
            }

            // 统计信息
            long bindingCount = chartInfos.stream().filter(c -> c.getBinding() != null).count();
            System.out.println("总图表数: " + chartInfos.size());
            System.out.println("有绑定数据的图表数: " + bindingCount);

        } catch (Exception e) {
            System.err.println("❌ 读取失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}