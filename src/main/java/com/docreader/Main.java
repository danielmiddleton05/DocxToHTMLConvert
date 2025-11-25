package com.docreader;

import com.docreader.converter.DocxToHtmlConverter;

import java.io.IOException;

/**
 * Main application class demonstrating how to use the DOCX to HTML converter.
 */
public class Main {

    public static void main(String[] args) {
        if (args.length < 2) {
            System.out.println("Usage: java -jar docx-to-html-converter.jar <input.docx> <output.html>");
            System.out.println("\nExample:");
            System.out.println("  java -jar docx-to-html-converter.jar document.docx document.html");
            System.exit(1);
        }

        String inputDocxPath = args[0];
        String outputHtmlPath = args[1];

        System.out.println("Converting DOCX to HTML...");
        System.out.println("Input:  " + inputDocxPath);
        System.out.println("Output: " + outputHtmlPath);

        try {
            DocxToHtmlConverter converter = new DocxToHtmlConverter();
            converter.convert(inputDocxPath, outputHtmlPath);
            
            System.out.println("\nConversion completed successfully!");
            System.out.println("HTML file saved to: " + outputHtmlPath);
            
        } catch (IOException e) {
            System.err.println("\nError during conversion: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
}
