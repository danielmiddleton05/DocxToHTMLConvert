package com.docreader.converter;

import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;
public class DocxToHtmlConverter {

    private StringBuilder htmlBuilder;

    public DocxToHtmlConverter() {
        this.htmlBuilder = new StringBuilder();
    }

    public void convert(String docxFilePath, String htmlFilePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(docxFilePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            initializeHtml();
            processDocument(document);
            finalizeHtml();

            writeToFile(htmlFilePath);
        }
    }

    public String convertToString(String docxFilePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(docxFilePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            initializeHtml();
            processDocument(document);
            finalizeHtml();

            return htmlBuilder.toString();
        }
    }

    private void initializeHtml() {
        htmlBuilder = new StringBuilder();
        htmlBuilder.append("<!DOCTYPE html>\n");
        htmlBuilder.append("<html lang=\"en\">\n");
        htmlBuilder.append("<head>\n");
        htmlBuilder.append("    <meta charset=\"UTF-8\">\n");
        htmlBuilder.append("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n");
        htmlBuilder.append("    <title>Converted Document</title>\n");
        htmlBuilder.append("    <style>\n");
        htmlBuilder.append("        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 20px; }\n");
        htmlBuilder.append("        table { border-collapse: collapse; width: 100%; margin: 20px 0; }\n");
        htmlBuilder.append("        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }\n");
        htmlBuilder.append("        th { background-color: #f2f2f2; font-weight: bold; }\n");
        htmlBuilder.append("        h1, h2, h3, h4, h5, h6 { margin-top: 20px; margin-bottom: 10px; }\n");
        htmlBuilder.append("        p { margin: 10px 0; }\n");
        htmlBuilder.append("        .bold { font-weight: bold; }\n");
        htmlBuilder.append("        .italic { font-style: italic; }\n");
        htmlBuilder.append("        .underline { text-decoration: underline; }\n");
        htmlBuilder.append("    </style>\n");
        htmlBuilder.append("</head>\n");
        htmlBuilder.append("<body>\n\n");
    }

    private void processDocument(XWPFDocument document) {
        List<IBodyElement> elements = document.getBodyElements();

        for (IBodyElement element : elements) {
            if (element instanceof XWPFParagraph) {
                processParagraph((XWPFParagraph) element);
            } else if (element instanceof XWPFTable) {
                processTable((XWPFTable) element);
            }
        }
    }

    private void processParagraph(XWPFParagraph paragraph) {
        String text = paragraph.getText();
        
        if (text == null || text.trim().isEmpty()) {
            return;
        }

        String style = paragraph.getStyle();
        
        if (style != null && style.toLowerCase().startsWith("heading")) {
            processHeading(paragraph, style);
        } else if (style != null && (style.equals("Heading1") || style.equals("Heading2") || style.equals("Heading3") || style.equals("Heading4") || style.equals("Heading5") || style.equals("Heading6"))) {
            processHeading(paragraph, style);
        } else {
            int outlineLevel = paragraph.getNumIlvl() != null ? paragraph.getNumIlvl().intValue() : -1;
            if (outlineLevel >= 0 && outlineLevel <= 8) {
                String headingTag = "h" + (outlineLevel + 1);
                htmlBuilder.append("    <").append(headingTag).append(">");
                processRuns(paragraph);
                htmlBuilder.append("</").append(headingTag).append(">\n");
            } else {
                htmlBuilder.append("    <p>");
                processRuns(paragraph);
                htmlBuilder.append("</p>\n");
            }
        }
    }

    private void processHeading(XWPFParagraph paragraph, String style) {
        int level = extractHeadingLevel(style);
        String headingTag = "h" + level;
        
        htmlBuilder.append("    <").append(headingTag).append(">");
        processRuns(paragraph);
        htmlBuilder.append("</").append(headingTag).append(">\n");
    }

    private int extractHeadingLevel(String style) {
        if (style == null) return 1;
        
        style = style.toLowerCase();
        
        if (style.contains("heading")) {
            for (int i = 1; i <= 9; i++) {
                if (style.contains(String.valueOf(i))) {
                    return Math.min(i, 6); // HTML only supports h1-h6
                }
            }
        }
        
        return 1; // Default to h1
    }

    private void processRuns(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        
        for (XWPFRun run : runs) {
            String text = run.getText(0);
            if (text == null || text.isEmpty()) {
                continue;
            }
            
            // Remove text enclosed in {{{ }}} (text to ignore)
            text = removeIgnoredText(text);
            if (text.isEmpty()) {
                continue;
            }
            
            boolean isBold = run.isBold();
            boolean isItalic = run.isItalic();
            UnderlinePatterns underline = run.getUnderline();
            boolean isUnderlined = underline != null && underline != UnderlinePatterns.NONE;
            
            if (isBold) {
                htmlBuilder.append("<strong>");
            }
            if (isItalic) {
                htmlBuilder.append("<em>");
            }
            if (isUnderlined) {
                htmlBuilder.append("<u>");
            }
            
            htmlBuilder.append(escapeHtml(text));
            
            if (isUnderlined) {
                htmlBuilder.append("</u>");
            }
            if (isItalic) {
                htmlBuilder.append("</em>");
            }
            if (isBold) {
                htmlBuilder.append("</strong>");
            }
        }
    }

    private String removeIgnoredText(String text) {
        if (text == null) {
            return "";
        }
        // Remove all text enclosed in {{{ }}}
        return text.replaceAll("\\{\\{\\{[^}]*\\}\\}\\}", "");
    }

    private void processTable(XWPFTable table) {
        htmlBuilder.append("    <table>\n");
        
        List<XWPFTableRow> rows = table.getRows();
        boolean isFirstRow = true;
        
        for (XWPFTableRow row : rows) {
            htmlBuilder.append("        <tr>\n");
            
            List<XWPFTableCell> cells = row.getTableCells();
            
            for (XWPFTableCell cell : cells) {
                String cellTag = isFirstRow ? "th" : "td";
                htmlBuilder.append("            <").append(cellTag).append(">");
                
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                for (int i = 0; i < paragraphs.size(); i++) {
                    XWPFParagraph paragraph = paragraphs.get(i);
                    String text = paragraph.getText();
                    
                    if (text != null && !text.trim().isEmpty()) {
                        processRuns(paragraph);
                        
                        if (i < paragraphs.size() - 1) {
                            htmlBuilder.append("<br>");
                        }
                    }
                }
                
                htmlBuilder.append("</").append(cellTag).append(">\n");
            }
            
            htmlBuilder.append("        </tr>\n");
            isFirstRow = false;
        }
        
        htmlBuilder.append("    </table>\n\n");
    }

    private String escapeHtml(String text) {
        if (text == null) {
            return "";
        }
        
        return text.replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;")
                   .replace("\"", "&quot;")
                   .replace("'", "&#39;");
    }

    private void finalizeHtml() {
        htmlBuilder.append("\n</body>\n");
        htmlBuilder.append("</html>");
    }

    private void writeToFile(String htmlFilePath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(htmlFilePath))) {
            writer.write(htmlBuilder.toString());
        }
    }

    public String getHtml() {
        return htmlBuilder.toString();
    }
}