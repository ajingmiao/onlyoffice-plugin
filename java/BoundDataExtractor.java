import java.io.*;
import java.util.zip.*;
import java.util.regex.*;
import java.util.*;

public class BoundDataExtractor {
    public static void main(String[] args) {
        String filePath = "demo.docx";

        System.out.println("=== OnlyOffice 图表绑定数据完整提取器 ===");
        System.out.println("文档: " + filePath);
        System.out.println();

        try {
            File file = new File(filePath);
            if (!file.exists()) {
                System.out.println("❌ 文件不存在: " + filePath);
                return;
            }

            try (ZipFile zipFile = new ZipFile(file)) {
                extractBoundData(zipFile);
            }

        } catch (Exception e) {
            System.err.println("❌ 分析失败: " + e.getMessage());
            e.printStackTrace();
        }
    }

    static void extractBoundData(ZipFile zipFile) throws IOException {
        ZipEntry customPropsEntry = zipFile.getEntry("docProps/custom.xml");
        if (customPropsEntry == null) {
            System.out.println("❌ 未找到自定义属性文件");
            return;
        }

        try (InputStream is = zipFile.getInputStream(customPropsEntry);
             BufferedReader reader = new BufferedReader(new InputStreamReader(is, "UTF-8"))) {

            String content = reader.lines()
                .collect(java.util.stream.Collectors.joining("\n"));

            // 查找所有chart-binding属性
            Pattern pattern = Pattern.compile("name=\"chart-binding:([^\"]+)\"[^>]*><vt:lpwstr>([^<]+)</vt:lpwstr>");
            Matcher matcher = pattern.matcher(content);

            int chartIndex = 1;
            while (matcher.find()) {
                String fingerprint = matcher.group(1);
                String jsonData = matcher.group(2);

                System.out.println("📊 图表 " + chartIndex + ":");
                System.out.println("   🔖 指纹: " + fingerprint);

                analyzeFingerprint(fingerprint);
                extractFullBoundData(jsonData);

                System.out.println("   " + "=".repeat(80));
                System.out.println();
                chartIndex++;
            }
        }
    }

    static void analyzeFingerprint(String fingerprint) {
        System.out.println("   🔍 指纹解析:");
        String[] parts = fingerprint.split("\\|");
        for (String part : parts) {
            if (part.startsWith("type:")) {
                System.out.println("      📊 图表类型: " + part.substring(5));
            } else if (part.startsWith("size:")) {
                System.out.println("      📏 尺寸: " + part.substring(5));
            } else if (part.startsWith("content:")) {
                System.out.println("      🔣 内容哈希: " + part.substring(8));
            } else if (part.startsWith("ctx:")) {
                System.out.println("      📋 上下文: " + part.substring(4) + " 个图表");
            }
        }
    }

    static void extractFullBoundData(String jsonData) {
        // 处理HTML实体编码
        String json = jsonData.replace("&quot;", "\"")
                            .replace("&amp;", "&")
                            .replace("&lt;", "<")
                            .replace("&gt;", ">")
                            .replace("\\\"", "\"");

        System.out.println("   📊 基本信息:");

        // 提取基本信息
        String chartId = extractJsonValue(json, "chartId");
        String chartType = extractJsonValue(json, "chartType");
        String createdAt = extractJsonValue(json, "createdAt");
        String lastUpdated = extractJsonValue(json, "lastUpdated");
        String fingerprint = extractJsonValue(json, "fingerprint");

        if (chartId != null) {
            System.out.println("      🆔 图表ID: " + chartId);
        }
        if (chartType != null) {
            System.out.println("      📊 图表类型: " + chartType);
        }
        if (createdAt != null) {
            System.out.println("      📅 创建时间: " + createdAt);
        }
        if (lastUpdated != null) {
            System.out.println("      🔄 更新时间: " + lastUpdated);
        }
        if (fingerprint != null) {
            System.out.println("      🔖 存储指纹: " + fingerprint);
        }

        // 检查是否有绑定数据
        if (json.contains("\"boundData\":")) {
            System.out.println("   ✅ 绑定状态: 已绑定数据");
            System.out.println("   📦 boundData 完整内容:");

            // 提取完整的boundData对象
            String boundDataJson = extractBoundDataObject(json);
            if (boundDataJson != null) {
                parseBoundDataObject(boundDataJson);
            }

            // 提取tagData对象
            if (json.contains("\"tagData\":")) {
                System.out.println("   🏷️ tagData 完整内容:");
                String tagDataJson = extractTagDataObject(json);
                if (tagDataJson != null) {
                    parseTagDataObject(tagDataJson);
                }
            }

            // 提取元数据
            if (json.contains("\"metadata\":")) {
                System.out.println("   📋 metadata 完整内容:");
                String metadataJson = extractMetadataObject(json);
                if (metadataJson != null) {
                    parseMetadataObject(metadataJson);
                }
            }

            // 显示原始JSON的一部分用于验证
            System.out.println("   📄 原始JSON预览 (前800字符):");
            String preview = json.length() > 800 ? json.substring(0, 800) + "..." : json;
            System.out.println("      " + preview.replace("\n", "\\n"));

        } else {
            System.out.println("   ❌ 绑定状态: 无绑定数据");
            System.out.println("   📄 完整JSON内容:");
            System.out.println("      " + json);
        }
    }

    static String extractBoundDataObject(String json) {
        return extractNestedObject(json, "\"boundData\":");
    }

    static String extractTagDataObject(String json) {
        return extractNestedObject(json, "\"tagData\":");
    }

    static String extractMetadataObject(String json) {
        return extractNestedObject(json, "\"metadata\":");
    }

    static String extractNestedObject(String json, String key) {
        int start = json.indexOf(key);
        if (start == -1) return null;

        int braceStart = json.indexOf("{", start);
        if (braceStart == -1) return null;

        int braceCount = 0;
        int end = braceStart;
        boolean inString = false;
        boolean escaped = false;

        for (int i = braceStart; i < json.length(); i++) {
            char c = json.charAt(i);

            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == '"') {
                inString = !inString;
                continue;
            }

            if (!inString) {
                if (c == '{') {
                    braceCount++;
                } else if (c == '}') {
                    braceCount--;
                    if (braceCount == 0) {
                        end = i;
                        break;
                    }
                }
            }
        }

        if (end > braceStart) {
            return json.substring(braceStart, end + 1);
        }
        return null;
    }

    static void parseBoundDataObject(String boundDataJson) {
        System.out.println("      📄 type: " + extractJsonValue(boundDataJson, "type"));
        System.out.println("      🏷️ rid: " + extractJsonValue(boundDataJson, "rid"));
        System.out.println("      📊 chartType: " + extractJsonValue(boundDataJson, "chartType"));
        System.out.println("      📄 originalFormat: " + extractJsonValue(boundDataJson, "originalFormat"));

        String tag = extractJsonValue(boundDataJson, "tag");
        if (tag != null) {
            System.out.println("      🏷️ tag (JSON字符串): " + tag);
            // 解析tag中的JSON
            parseTagString(tag);
        }

        String chartId = extractJsonValue(boundDataJson, "chartId");
        if (chartId != null) {
            System.out.println("      🆔 chartId: " + chartId);
        }

        System.out.println("      📄 完整boundData JSON:");
        System.out.println("         " + boundDataJson);
    }

    static void parseTagDataObject(String tagDataJson) {
        System.out.println("      ⏰ _t: " + extractJsonValue(tagDataJson, "_t"));
        System.out.println("      🔍 trackId: " + extractJsonValue(tagDataJson, "trackId"));

        // 提取groupFields数组
        if (tagDataJson.contains("\"groupFields\":[")) {
            int start = tagDataJson.indexOf("\"groupFields\":[");
            int end = tagDataJson.indexOf("]", start);
            if (end != -1) {
                String groupFieldsStr = tagDataJson.substring(start + 15, end);
                System.out.println("      📊 groupFields: [" + groupFieldsStr + "]");
            }
        }

        System.out.println("      📄 完整tagData JSON:");
        System.out.println("         " + tagDataJson);
    }

    static void parseMetadataObject(String metadataJson) {
        System.out.println("      📄 sourceFormat: " + extractJsonValue(metadataJson, "sourceFormat"));
        System.out.println("      🔄 convertedAt: " + extractJsonValue(metadataJson, "convertedAt"));
        System.out.println("      🔍 trackId: " + extractJsonValue(metadataJson, "trackId"));
        System.out.println("      ⏰ timestamp: " + extractJsonValue(metadataJson, "timestamp"));

        // 提取groupFields数组
        if (metadataJson.contains("\"groupFields\":[")) {
            int start = metadataJson.indexOf("\"groupFields\":[");
            int end = metadataJson.indexOf("]", start);
            if (end != -1) {
                String groupFieldsStr = metadataJson.substring(start + 15, end);
                System.out.println("      📊 groupFields: [" + groupFieldsStr + "]");
            }
        }

        System.out.println("      📄 完整metadata JSON:");
        System.out.println("         " + metadataJson);
    }

    static void parseTagString(String tag) {
        try {
            // 处理转义的JSON字符串
            tag = tag.replace("\\\"", "\"");

            if (tag.startsWith("{") && tag.endsWith("}")) {
                System.out.println("         🔍 解析tag内容:");
                System.out.println("            ⏰ _t: " + extractJsonValue(tag, "_t"));
                System.out.println("            🔍 trackId: " + extractJsonValue(tag, "trackId"));

                // 提取groupFields数组
                if (tag.contains("\"groupFields\":[")) {
                    int start = tag.indexOf("\"groupFields\":[");
                    int end = tag.indexOf("]", start);
                    if (end != -1) {
                        String groupFieldsStr = tag.substring(start + 15, end);
                        System.out.println("            📊 groupFields: [" + groupFieldsStr + "]");
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("         ⚠️ tag解析失败: " + e.getMessage());
        }
    }

    static String extractJsonValue(String json, String key) {
        Pattern pattern = Pattern.compile("\"" + key + "\"\\s*:\\s*\"([^\"]+)\"");
        Matcher matcher = pattern.matcher(json);
        if (matcher.find()) {
            return matcher.group(1);
        }

        // 尝试匹配数字值
        Pattern numberPattern = Pattern.compile("\"" + key + "\"\\s*:\\s*([0-9]+)");
        Matcher numberMatcher = numberPattern.matcher(json);
        if (numberMatcher.find()) {
            return numberMatcher.group(1);
        }

        return null;
    }
}