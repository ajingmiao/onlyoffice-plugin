import java.io.*;
import java.util.zip.*;
import java.util.regex.*;
import java.util.*;

public class DebugBoundDataExtractor {
    public static void main(String[] args) {
        System.out.println("=== DEBUG: OnlyOffice Chart Binding Data Extractor ===");

        String filePath = "demo.docx";
        System.out.println("DEBUG: Looking for file: " + filePath);

        try {
            File file = new File(filePath);
            System.out.println("DEBUG: File object created");
            System.out.println("DEBUG: File exists: " + file.exists());
            System.out.println("DEBUG: File absolute path: " + file.getAbsolutePath());

            if (!file.exists()) {
                System.out.println("ERROR: File not found: " + filePath);
                System.out.println("DEBUG: Current working directory: " + System.getProperty("user.dir"));
                System.out.println("DEBUG: Please make sure demo.docx is in the current directory");
                return;
            }

            System.out.println("DEBUG: File size: " + file.length() + " bytes");
            System.out.println("DEBUG: Opening ZIP file...");

            try (ZipFile zipFile = new ZipFile(file)) {
                System.out.println("DEBUG: ZIP file opened successfully");
                extractBoundDataDebug(zipFile);
            }

        } catch (Exception e) {
            System.err.println("ERROR: Analysis failed: " + e.getMessage());
            System.err.println("DEBUG: Stack trace:");
            e.printStackTrace();
        }
    }

    static void extractBoundDataDebug(ZipFile zipFile) throws IOException {
        System.out.println("DEBUG: Looking for docProps/custom.xml...");

        ZipEntry customPropsEntry = zipFile.getEntry("docProps/custom.xml");
        if (customPropsEntry == null) {
            System.out.println("ERROR: docProps/custom.xml not found");
            System.out.println("DEBUG: Available entries in ZIP:");
            zipFile.entries().asIterator().forEachRemaining(entry -> {
                System.out.println("  - " + entry.getName());
            });
            return;
        }

        System.out.println("DEBUG: Found docProps/custom.xml, size: " + customPropsEntry.getSize() + " bytes");

        try (InputStream is = zipFile.getInputStream(customPropsEntry);
             BufferedReader reader = new BufferedReader(new InputStreamReader(is, "UTF-8"))) {

            System.out.println("DEBUG: Reading custom properties...");
            String content = reader.lines()
                .collect(java.util.stream.Collectors.joining("\n"));

            System.out.println("DEBUG: Custom properties content length: " + content.length() + " characters");
            System.out.println("DEBUG: First 200 characters: " + content.substring(0, Math.min(200, content.length())));

            // Find all chart-binding properties
            Pattern pattern = Pattern.compile("name=\"chart-binding:([^\"]+)\"[^>]*><vt:lpwstr>([^<]+)</vt:lpwstr>");
            Matcher matcher = pattern.matcher(content);

            System.out.println("DEBUG: Searching for chart-binding patterns...");

            int chartIndex = 1;
            boolean foundAny = false;

            while (matcher.find()) {
                foundAny = true;
                String fingerprint = matcher.group(1);
                String jsonData = matcher.group(2);

                System.out.println("\n=== CHART " + chartIndex + " ===");
                System.out.println("Fingerprint: " + fingerprint);

                analyzeFingerprint(fingerprint);
                extractFullBoundDataDebug(jsonData);

                System.out.println("=====================================");
                chartIndex++;
            }

            if (!foundAny) {
                System.out.println("DEBUG: No chart-binding patterns found");
                System.out.println("DEBUG: Checking if content contains 'chart-binding'...");
                if (content.contains("chart-binding")) {
                    System.out.println("DEBUG: Found 'chart-binding' text in content");
                    int index = content.indexOf("chart-binding");
                    int start = Math.max(0, index - 100);
                    int end = Math.min(content.length(), index + 200);
                    System.out.println("DEBUG: Context around 'chart-binding':");
                    System.out.println(content.substring(start, end));
                } else {
                    System.out.println("DEBUG: No 'chart-binding' text found in content");
                }
            } else {
                System.out.println("\nDEBUG: Found " + (chartIndex - 1) + " chart bindings total");
            }
        }
    }

    static void analyzeFingerprint(String fingerprint) {
        System.out.println("Fingerprint Analysis:");
        String[] parts = fingerprint.split("\\|");
        for (String part : parts) {
            if (part.startsWith("type:")) {
                System.out.println("  Chart Type: " + part.substring(5));
            } else if (part.startsWith("size:")) {
                System.out.println("  Size: " + part.substring(5));
            } else if (part.startsWith("content:")) {
                System.out.println("  Content Hash: " + part.substring(8));
            } else if (part.startsWith("ctx:")) {
                System.out.println("  Context: " + part.substring(4) + " charts");
            } else {
                System.out.println("  Other: " + part);
            }
        }
    }

    static void extractFullBoundDataDebug(String jsonData) {
        // Handle HTML entities
        String json = jsonData.replace("&quot;", "\"")
                            .replace("&amp;", "&")
                            .replace("&lt;", "<")
                            .replace("&gt;", ">")
                            .replace("\\\"", "\"");

        System.out.println("Basic Info:");

        // Extract basic information
        String chartId = extractJsonValue(json, "chartId");
        String chartType = extractJsonValue(json, "chartType");
        String createdAt = extractJsonValue(json, "createdAt");
        String lastUpdated = extractJsonValue(json, "lastUpdated");

        System.out.println("  Chart ID: " + (chartId != null ? chartId : "NOT FOUND"));
        System.out.println("  Chart Type: " + (chartType != null ? chartType : "NOT FOUND"));
        System.out.println("  Created At: " + (createdAt != null ? createdAt : "NOT FOUND"));
        System.out.println("  Last Updated: " + (lastUpdated != null ? lastUpdated : "NOT FOUND"));

        // Check for bound data
        if (json.contains("\"boundData\":")) {
            System.out.println("  Binding Status: HAS BOUND DATA");
            System.out.println("boundData Details:");

            // Extract complete boundData object
            String boundDataJson = extractBoundDataObject(json);
            if (boundDataJson != null) {
                parseBoundDataObjectDebug(boundDataJson);
            } else {
                System.out.println("  ERROR: Could not extract boundData object");
            }

        } else {
            System.out.println("  Binding Status: NO BOUND DATA");
            System.out.println("Complete JSON Content:");
            System.out.println("  " + (json.length() > 300 ? json.substring(0, 300) + "..." : json));
        }
    }

    static String extractBoundDataObject(String json) {
        int start = json.indexOf("\"boundData\":");
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

    static void parseBoundDataObjectDebug(String boundDataJson) {
        System.out.println("  type: " + extractJsonValue(boundDataJson, "type"));
        System.out.println("  rid: " + extractJsonValue(boundDataJson, "rid"));
        System.out.println("  chartType: " + extractJsonValue(boundDataJson, "chartType"));
        System.out.println("  originalFormat: " + extractJsonValue(boundDataJson, "originalFormat"));

        String tag = extractJsonValue(boundDataJson, "tag");
        if (tag != null) {
            System.out.println("  tag (JSON string): " + (tag.length() > 100 ? tag.substring(0, 100) + "..." : tag));
        }

        System.out.println("Complete boundData JSON:");
        System.out.println("  " + boundDataJson);
    }

    static String extractJsonValue(String json, String key) {
        Pattern pattern = Pattern.compile("\"" + key + "\"\\s*:\\s*\"([^\"]+)\"");
        Matcher matcher = pattern.matcher(json);
        if (matcher.find()) {
            return matcher.group(1);
        }

        // Try to match numeric values
        Pattern numberPattern = Pattern.compile("\"" + key + "\"\\s*:\\s*([0-9]+)");
        Matcher numberMatcher = numberPattern.matcher(json);
        if (numberMatcher.find()) {
            return numberMatcher.group(1);
        }

        return null;
    }
}