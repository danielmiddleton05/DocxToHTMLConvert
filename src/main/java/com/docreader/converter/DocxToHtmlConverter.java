package com.docreader.converter;

import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;

/**
 * Converts DOCX files to HTML format.
 * Handles headers, paragraphs, tables, and basic text formatting.
 */
public class DocxToHtmlConverter {

    private StringBuilder htmlBuilder;

    public DocxToHtmlConverter() {
        this.htmlBuilder = new StringBuilder();
    }

    /**
     * Converts a DOCX file to HTML format.
     *
     * @param docxFilePath Path to the input DOCX file
     * @param htmlFilePath Path to the output HTML file
     * @throws IOException If file reading/writing fails
     */
    public void convert(String docxFilePath, String htmlFilePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(docxFilePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            initializeHtml();
            processDocument(document);
            finalizeHtml();

            writeToFile(htmlFilePath);
        }
    }

    /**
     * Converts a DOCX file and returns the HTML as a String.
     *
     * @param docxFilePath Path to the input DOCX file
     * @return HTML content as String
     * @throws IOException If file reading fails
     */
    public String convertToString(String docxFilePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(docxFilePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            initializeHtml();
            processDocument(document);
            finalizeHtml();

            return htmlBuilder.toString();
        }
    }

    /**
     * Initializes the HTML document structure.
     */
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

    /**
     * Processes the entire DOCX document.
     *
     * @param document The XWPFDocument to process
     */
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

    /**
     * Processes a paragraph, converting it to the appropriate HTML heading or paragraph tag.
     *
     * @param paragraph The XWPFParagraph to process
     */
    private void processParagraph(XWPFParagraph paragraph) {
        String text = paragraph.getText();
        
        // Skip empty paragraphs
        if (text == null || text.trim().isEmpty()) {
            return;
        }

        String style = paragraph.getStyle();
        
        // Check if paragraph is a heading
        if (style != null && style.toLowerCase().startsWith("heading")) {
            processHeading(paragraph, style);
        } else if (style != null && (style.equals("Heading1") || style.equals("Heading2") || 
                   style.equals("Heading3") || style.equals("Heading4") || 
                   style.equals("Heading5") || style.equals("Heading6"))) {
            processHeading(paragraph, style);
        } else {
            // Check paragraph outline level for headings
            int outlineLevel = paragraph.getNumIlvl() != null ? paragraph.getNumIlvl().intValue() : -1;
            if (outlineLevel >= 0 && outlineLevel <= 8) {
                String headingTag = "h" + (outlineLevel + 1);
                htmlBuilder.append("    <").append(headingTag).append(">");
                processRuns(paragraph);
                htmlBuilder.append("</").append(headingTag).append(">\n");
            } else {
                // Regular paragraph
                htmlBuilder.append("    <p>");
                processRuns(paragraph);
                htmlBuilder.append("</p>\n");
            }
        }
    }

    /**
     * Processes a heading paragraph.
     *
     * @param paragraph The heading paragraph
     * @param style The style name
     */
    private void processHeading(XWPFParagraph paragraph, String style) {
        int level = extractHeadingLevel(style);
        String headingTag = "h" + level;
        
        htmlBuilder.append("    <").append(headingTag).append(">");
        processRuns(paragraph);
        htmlBuilder.append("</").append(headingTag).append(">\n");
    }

    /**
     * Extracts the heading level from the style name.
     *
     * @param style The style name
     * @return The heading level (1-6)
     */
    private int extractHeadingLevel(String style) {
        if (style == null) return 1;
        
        style = style.toLowerCase();
        
        // Try to extract number from style name
        if (style.contains("heading")) {
            for (int i = 1; i <= 9; i++) {
                if (style.contains(String.valueOf(i))) {
                    return Math.min(i, 6); // HTML only supports h1-h6
                }
            }
        }
        
        return 1; // Default to h1
    }

    /**
     * Processes the runs within a paragraph, applying text formatting.
     *
     * @param paragraph The paragraph containing the runs
     */
    private void processRuns(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        
        for (XWPFRun run : runs) {
            String text = run.getText(0);
            if (text == null || text.isEmpty()) {
                continue;
            }
            
            boolean isBold = run.isBold();
            boolean isItalic = run.isItalic();
            UnderlinePatterns underline = run.getUnderline();
            boolean isUnderlined = underline != null && underline != UnderlinePatterns.NONE;
            
            // Apply formatting
            if (isBold) {
                htmlBuilder.append("<strong>");
            }
            if (isItalic) {
                htmlBuilder.append("<em>");
            }
            if (isUnderlined) {
                htmlBuilder.append("<u>");
            }
            
            // Escape HTML special characters
            htmlBuilder.append(escapeHtml(text));
            
            // Close formatting tags in reverse order
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

    /**
     * Processes a table, converting it to HTML table format.
     *
     * @param table The XWPFTable to process
     */
    private void processTable(XWPFTable table) {
        htmlBuilder.append("    <table>\n");
        
        List<XWPFTableRow> rows = table.getRows();
        boolean isFirstRow = true;
        
        for (XWPFTableRow row : rows) {
            htmlBuilder.append("        <tr>\n");
            
            List<XWPFTableCell> cells = row.getTableCells();
            
            for (XWPFTableCell cell : cells) {
                // Use th for first row (header), td for others
                String cellTag = isFirstRow ? "th" : "td";
                htmlBuilder.append("            <").append(cellTag).append(">");
                
                // Process all paragraphs in the cell
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                for (int i = 0; i < paragraphs.size(); i++) {
                    XWPFParagraph paragraph = paragraphs.get(i);
                    String text = paragraph.getText();
                    
                    if (text != null && !text.trim().isEmpty()) {
                        processRuns(paragraph);
                        
                        // Add line break between paragraphs (except last one)
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

    /**
     * Escapes HTML special characters.
     *
     * @param text The text to escape
     * @return The escaped text
     */
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

    /**
     * Finalizes the HTML document structure.
     */
    private void finalizeHtml() {
        htmlBuilder.append("\n</body>\n");
        htmlBuilder.append("</html>");
    }

    /**
     * Writes the HTML content to a file.
     *
     * @param htmlFilePath Path to the output file
     * @throws IOException If file writing fails
     */
    private void writeToFile(String htmlFilePath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(htmlFilePath))) {
            writer.write(htmlBuilder.toString());
        }
    }

    /**
     * Gets the generated HTML as a String.
     *
     * @return The HTML content
     */
    public String getHtml() {
        return htmlBuilder.toString();
    }
}
