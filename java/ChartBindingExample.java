package com.onlyoffice.chart;

import java.util.List;

/**
 * Java POI图表绑定读取器使用示例
 */
public class ChartBindingExample {

    public static void main(String[] args) {
        // 示例1: 读取指定Word文档中的所有图表绑定
        String filePath = "path/to/your/document.docx";

        ChartBindingReader reader = new ChartBindingReader();

        try {
            System.out.println("=== 示例1: 读取所有图表绑定 ===");

            List<ChartBindingReader.ChartInfo> chartInfos = reader.readChartBindings(filePath);

            for (ChartBindingReader.ChartInfo chartInfo : chartInfos) {
                System.out.println("图表索引: " + chartInfo.getIndex());
                System.out.println("图表类型: " + chartInfo.getChartType());
                System.out.println("指纹: " + chartInfo.getFingerprint());

                if (chartInfo.getBinding() != null) {
                    ChartBindingReader.ChartBinding binding = chartInfo.getBinding();
                    System.out.println("绑定数据:");
                    System.out.println("  - Chart ID: " + binding.getChartId());
                    System.out.println("  - Chart Type: " + binding.getChartType());
                    System.out.println("  - RID: " + binding.getRid());
                    System.out.println("  - Tag: " + binding.getTag());
                    System.out.println("  - Created At: " + binding.getCreatedAt());

                    if (binding.getBoundData() != null) {
                        System.out.println("  - Bound Data: " + binding.getBoundData());
                    }

                    if (binding.getTagData() != null) {
                        System.out.println("  - Tag Data: " + binding.getTagData());
                    }
                } else {
                    System.out.println("无绑定数据");
                }
                System.out.println("---");
            }

            // 示例2: 根据chartId查找特定图表绑定
            System.out.println("\n=== 示例2: 根据chartId查找绑定 ===");
            String targetChartId = "chart-idx_0_id_123_type_bar-chart_1a2b3c";

            ChartBindingReader.ChartBinding specificBinding = reader.findBindingByChartId(filePath, targetChartId);
            if (specificBinding != null) {
                System.out.println("找到指定图表绑定:");
                System.out.println("Chart ID: " + specificBinding.getChartId());
                System.out.println("Chart Type: " + specificBinding.getChartType());
                System.out.println("RID: " + specificBinding.getRid());
                System.out.println("Tag: " + specificBinding.getTag());
            } else {
                System.out.println("未找到chartId为 '" + targetChartId + "' 的图表绑定");
            }

            // 示例3: 获取所有有绑定数据的图表
            System.out.println("\n=== 示例3: 获取所有绑定数据 ===");
            List<ChartBindingReader.ChartBinding> allBindings = reader.getAllBindings(filePath);

            System.out.println("文档中共有 " + allBindings.size() + " 个图表绑定:");
            for (ChartBindingReader.ChartBinding binding : allBindings) {
                System.out.println("- " + binding.getChartId() + " (" + binding.getChartType() + ")");
                if (binding.getRid() != null) {
                    System.out.println("  RID: " + binding.getRid());
                }
                if (binding.getTag() != null) {
                    System.out.println("  Tag: " + binding.getTag());
                }
            }

        } catch (Exception e) {
            System.err.println("读取失败: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 演示如何在Spring Boot应用中使用
     */
    public static class SpringBootExample {

        private ChartBindingReader chartReader = new ChartBindingReader();

        public List<ChartBindingReader.ChartBinding> getChartBindings(String documentPath) {
            try {
                List<ChartBindingReader.ChartInfo> chartInfos = chartReader.readChartBindings(documentPath);
                return chartInfos.stream()
                        .filter(info -> info.getBinding() != null)
                        .map(ChartBindingReader.ChartInfo::getBinding)
                        .collect(java.util.stream.Collectors.toList());
            } catch (Exception e) {
                throw new RuntimeException("Failed to read chart bindings", e);
            }
        }

        public ChartBindingReader.ChartBinding findChartByRid(String documentPath, String rid) {
            try {
                List<ChartBindingReader.ChartBinding> bindings = chartReader.getAllBindings(documentPath);
                return bindings.stream()
                        .filter(binding -> rid.equals(binding.getRid()))
                        .findFirst()
                        .orElse(null);
            } catch (Exception e) {
                throw new RuntimeException("Failed to find chart by RID", e);
            }
        }
    }
}