# 📄 DOCX Generator

A simple web app that converts your JavaScript code into Microsoft Word documents (DOCX). Just upload a `.js` file that creates a document using the `docx` library, and get your DOCX file instantly.

## 🚀 Try It Online

**[Live Demo](https://web-production-0f7e8.up.railway.app)**
**[Demo_video](https://drive.google.com/drive/folders/1yFsrbwSpv8quEBNlrtx4qRs3uSWynYyp?usp=drive_link)**



Simply upload your JavaScript file and download the generated DOCX!

## 🎯 How It Works

1. **Upload** your JavaScript file (`.js`)
2. **Server executes** your code in a secure sandbox
3. **Generates** a DOCX document using the docx library
4. **Downloads** your file automatically

## 📝 Code Requirements

Your JavaScript file must create a `Document` object and use `Packer.toBuffer()` to generate the DOCX:

```javascript
const { Document, Packer, Paragraph, TextRun } = require('docx');
const fs = require('fs');

const doc = new Document({
  sections: [{
    children: [
      new Paragraph({
        children: [new TextRun("Hello, World!")]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("output.docx", buffer);
});
```

## 🏃 Run Locally

### Prerequisites
- Node.js (v14+)
- npm (v6+)

### Installation

```bash
# Clone the repository
git clone https://github.com/yugbathla30/givememydocclaude.git
cd givememydocclaude

# Install dependencies
npm install

# Start the server
npm start
```

The app will be available at **http://localhost:3000**

### Environment Variables

Create a `.env` file in the root directory (optional):

```env
NODE_ENV=development
PORT=3000
UPLOAD_DIR=./uploads
OUTPUT_DIR=./output
```

See `.env.example` for all available options.

## 📦 Technologies

- **Express.js** - Web server
- **docx** - DOCX document generation
- **Node.js VM** - Secure code execution sandbox
- **Docker** - Containerization

## 🔒 Security

- Code runs in an isolated VM sandbox
- File writes are restricted to the output directory
- Only `docx` and `fs` modules are available
- 30-second execution timeout
- 5MB file size limit

## 📚 Supported DOCX Features

- Paragraphs, Headings, Text Formatting
- Tables with custom styling
- Headers & Footers
- Page breaks and page orientation
- Hyperlinks (external and internal)
- Footnotes
- Lists and numbering
- Images
- Table of Contents
- And more!

## 🐳 Docker

```bash
# Build and run with Docker
docker-compose up

# Or with Docker directly
docker build -t docx-generator .
docker run -p 3000:3000 docx-generator
```

## 📄 License

MIT

---

**Questions?** Check the [docx library docs](https://docx.js.org) for complete API reference.
