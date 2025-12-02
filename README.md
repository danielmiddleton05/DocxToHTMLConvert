# DOCX to HTML Converter

A simple Java application that converts Microsoft Word (.docx) files to clean, styled HTML. Perfect for converting documentation, reports, or any Word documents into web-ready HTML format.

## Quick Start

### What You Need

- **Java 11 or higher** installed on your computer
- **Maven 3.6 or higher** (for building the project)
- A Word document (.docx) you want to convert

---

## Setup Instructions

### Step 1: Get the Project

If you haven't already, clone or download this project to your computer.

### Step 2: Open Terminal/Command Prompt

Navigate to the project folder:

```powershell
cd path\to\DocReader
```

### Step 3: Build the Application

Run this command to compile and package the application:

```powershell
mvn clean package
```

This creates a runnable JAR file in the `target` folder.

**Setup Complete!** You're ready to convert documents.

---

## How to Use

### Simple 3-Step Process

#### **Step 1: Place Your Word Document**

Put your `.docx` file in the `input` folder:

```
DocReader/
├── input/           ← Put your .docx files here
│   └── ARN Build Doc.docx
├── output/          ← HTML files will appear here
└── ...
```

#### **Step 2: Run the Converter**

Use this command format:

```powershell
java -jar target\docx-to-html-converter-1.0-SNAPSHOT.jar input\YourFile.docx output\YourFile.html
```

**Real Example:**

```powershell
java -jar target\docx-to-html-converter-1.0-SNAPSHOT.jar "input\ARN Build Doc.docx" "output\ARN Build Doc.html"
```

#### **Step 3: Find Your HTML**

Check the `output` folder for your converted HTML file!

---

## Usage Examples

### Example 1: Basic Conversion

```powershell
java -jar target\docx-to-html-converter-1.0-SNAPSHOT.jar input\document.docx output\document.html
```

### Example 2: Using Full Paths

```powershell
java -jar target\docx-to-html-converter-1.0-SNAPSHOT.jar "C:\My Documents\report.docx" "C:\Web Files\report.html"
```

### Example 3: Files with Spaces in Names

Always use quotes for filenames with spaces:

```powershell
java -jar target\docx-to-html-converter-1.0-SNAPSHOT.jar "input\My Report 2025.docx" "output\My Report 2025.html"
```

---

## Features

### What Gets Converted?

#### **Headers**

- Heading 1 → `<h1>` (large header)
- Heading 2 → `<h2>` (medium header)
- Heading 3 → `<h3>` (smaller header)
- Up to Heading 6 → `<h6>`

#### **Text Formatting**

- **Bold text** → `<strong>`
- _Italic text_ → `<em>`
- <u>Underlined text</u> → `<u>`
- Combined formatting supported!

#### **Tables**

- Word tables convert to HTML tables
- First row becomes table headers
- Full structure preserved
- Styled with borders and spacing

#### **Ignored Text**

Text wrapped in `{{{triple curly braces}}}` will be automatically removed from the output.

- Example: `Hello {{{ignore this}}} World` → `Hello  World`

---

## Advanced Usage

### Use in Your Own Java Code

```java
import com.docreader.converter.DocxToHtmlConverter;
import java.io.IOException;

public class MyConverter {
    public static void main(String[] args) {
        try {
            DocxToHtmlConverter converter = new DocxToHtmlConverter();

            // Option 1: Convert and save to file
            converter.convert("input.docx", "output.html");

            // Option 2: Get HTML as a string
            String html = converter.convertToString("input.docx");
            System.out.println(html);

        } catch (IOException e) {
            System.err.println("Error: " + e.getMessage());
        }
    }
}
```

---

## Project Structure

```
DocReader/
├── input/                          ← Place .docx files here
├── output/                         ← HTML files appear here
├── src/
│   └── main/
│       └── java/
│           └── com/
│               └── docreader/
│                   ├── Main.java                    ← Entry point
│                   └── converter/
│                       └── DocxToHtmlConverter.java ← Conversion logic
├── target/
│   └── docx-to-html-converter-1.0-SNAPSHOT.jar     ← Built application
├── pom.xml                         ← Maven configuration
└── README.md                       ← This file
```

---

## Troubleshooting

### "Command not found: java"

- Java is not installed or not in your PATH
- Download Java from [adoptium.net](https://adoptium.net/)

### "Command not found: mvn"

- Maven is not installed or not in your PATH
- Download Maven from [maven.apache.org](https://maven.apache.org/)

### "File not found" Error

- Check that your file paths are correct
- Use quotes around filenames with spaces
- Use full paths if the file isn't in input/output folders

### Build Errors

Try cleaning and rebuilding:

```powershell
mvn clean
mvn package
```

---

## Technical Details

### Dependencies

- **Apache POI 5.2.5** - Reads Word documents
- **Apache POI OOXML** - Handles .docx format
- **XML Beans 5.1.1** - XML processing
- **Commons Collections 4.4** - Utilities

### Output HTML Structure

The converter generates standards-compliant HTML5 with:

- Proper DOCTYPE and meta tags
- Embedded CSS styling
- Semantic HTML elements
- Escaped special characters

---

## License

This project is open source and available for educational and commercial use.

---

## Contributing

Found a bug? Have a feature idea? Contributions welcome!

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

---

## Questions?

If you run into issues, check:

1. Java and Maven are properly installed
2. File paths are correct
3. Input file is a valid .docx document
4. You have write permissions to the output folder
