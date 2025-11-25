# DOCX to HTML Converter

A Java tool that converts Microsoft Word (DOCX) files to HTML format, preserving document structure including headers, tables, and text formatting.

## Features

- **Header Conversion**: Converts Word heading styles to HTML `<h1>`, `<h2>`, `<h3>`, etc.
- **Table Support**: Converts Word tables to HTML tables with proper structure
- **Text Formatting**: Preserves bold, italic, and underline formatting
- **Clean HTML Output**: Generates well-formatted HTML with CSS styling

## Requirements

- Java 11 or higher
- Maven 3.6 or higher

## Installation

1. Clone or download this project
2. Navigate to the project directory
3. Build the project using Maven:

```bash
mvn clean package
```

## Usage

### Command Line

Run the converter from the command line:

```bash
java -jar target/docx-to-html-converter-1.0-SNAPSHOT.jar input.docx output.html
```

### Programmatic Usage

You can also use the converter in your own Java code:

```java
import com.docreader.converter.DocxToHtmlConverter;

public class Example {
    public static void main(String[] args) {
        try {
            DocxToHtmlConverter converter = new DocxToHtmlConverter();
            
            // Convert and save to file
            converter.convert("input.docx", "output.html");
            
            // Or get HTML as string
            String html = converter.convertToString("input.docx");
            System.out.println(html);
            
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
```

## Supported Conversions

### Headers
- Word Heading 1 → `<h1>`
- Word Heading 2 → `<h2>`
- Word Heading 3 → `<h3>`
- And so on up to `<h6>`

### Text Formatting
- **Bold** → `<strong>`
- *Italic* → `<em>`
- Underline → `<u>`

### Tables
- Word tables → HTML `<table>` with proper rows and cells
- First row treated as header (`<th>`)
- Subsequent rows as data cells (`<td>`)

### Paragraphs
- Regular paragraphs → `<p>`
- Multiple paragraphs preserved

## Project Structure

```
DocReader/
├── pom.xml
├── README.md
└── src/
    └── main/
        └── java/
            └── com/
                └── docreader/
                    ├── Main.java
                    └── converter/
                        └── DocxToHtmlConverter.java
```

## Dependencies

- **Apache POI 5.2.5**: For reading and processing DOCX files
- **Apache POI OOXML**: For handling Office Open XML format
- **XML Beans**: For XML processing
- **Commons Collections**: Utility library

## Building from Source

```bash
# Clean and compile
mvn clean compile

# Run tests (if any)
mvn test

# Package as JAR
mvn package

# The JAR will be created in target/ directory
```

## Example

Given a Word document with:
- Heading 1: "My Document"
- Paragraph: "This is a sample paragraph with **bold** text."
- Table with 2 rows and 2 columns

The converter will produce:

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Converted Document</title>
    <style>
        /* CSS styling */
    </style>
</head>
<body>
    <h1>My Document</h1>
    <p>This is a sample paragraph with <strong>bold</strong> text.</p>
    <table>
        <tr>
            <th>Header 1</th>
            <th>Header 2</th>
        </tr>
        <tr>
            <td>Data 1</td>
            <td>Data 2</td>
        </tr>
    </table>
</body>
</html>
```

## License

This project is open source and available for educational and commercial use.

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests.
