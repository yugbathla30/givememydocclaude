const express = require('express');
const fileUpload = require('express-fileupload');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');
const vm = require('vm');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;
const NODE_ENV = process.env.NODE_ENV || 'development';
const MAX_FILE_SIZE = 5 * 1024 * 1024; // 5MB
const UPLOAD_DIR = process.env.UPLOAD_DIR || path.join(__dirname, 'uploads');
const OUTPUT_DIR = process.env.OUTPUT_DIR || path.join(__dirname, 'output');

// Middleware
app.use(express.static('public'));
app.use(fileUpload({ 
  limits: { fileSize: MAX_FILE_SIZE },
  useTempFiles: true,
  tempFileDir: '/tmp/',
  createParentPath: true
}));

// Create necessary directories
if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
}
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', environment: NODE_ENV });
});

// Cleanup function
function cleanupFiles(uploadPath, outputPath) {
  try {
    if (uploadPath && fs.existsSync(uploadPath)) {
      fs.unlinkSync(uploadPath);
    }
    if (outputPath && fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }
  } catch (err) {
    console.warn('Cleanup error:', err.message);
  }
}

// API endpoint to handle file upload and execution
app.post('/api/generate-docx', async (req, res) => {
  const fileId = uuidv4();
  const uploadPath = path.join(UPLOAD_DIR, `${fileId}.js`);
  const outputPath = path.join(OUTPUT_DIR, `${fileId}.docx`);

  try {
    if (!req.files || !req.files.jsFile) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const jsFile = req.files.jsFile;

    // Validate file type
    if (!jsFile.name.endsWith('.js')) {
      return res.status(400).json({ error: 'File must be a JavaScript (.js) file' });
    }

    // Validate file size
    if (jsFile.size > MAX_FILE_SIZE) {
      return res.status(400).json({ error: `File size exceeds ${MAX_FILE_SIZE / 1024 / 1024}MB limit` });
    }

    // Save uploaded file
    await jsFile.mv(uploadPath);

    // Read the uploaded JS file
    let jsCode = fs.readFileSync(uploadPath, 'utf-8');

    // Convert double-quoted strings containing literal newlines to template literals
    jsCode = jsCode.replace(/(")((?:[^"\\]|\\.|\n)*?)(")/g, (match, openQuote, content, closeQuote) => {
      if (content.includes('\n') && openQuote === '"') {
        const escaped = content
          .replace(/\\/g, '\\\\')
          .replace(/`/g, '\\`')
          .replace(/\$/g, '\\$');
        return '`' + escaped + '`';
      }
      return match;
    });
    
    // Don't remove Packer code - let it run and we'll catch the file write attempt
    // The user's fs.writeFileSync will be intercepted by our sandbox fs

    // Create a safe context with required modules
    // Importing ALL docx components mentioned in SKILL.md for comprehensive support
    const { 
      Document, Packer, Paragraph, TextRun, 
      Table, TableRow, TableCell, 
      ImageRun,
      Header, Footer,
      HeadingLevel, AlignmentType, LevelFormat, BorderStyle, 
      PageOrientation,
      ExternalHyperlink, InternalHyperlink, Bookmark, 
      FootnoteReferenceRun,
      PositionalTab, PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader,
      TabStopType, TabStopPosition, 
      Column, SectionType,
      TableOfContents,
      WidthType, ShadingType,
      VerticalAlign,
      PageNumber, PageBreak 
    } = require('docx');

    const sandbox = {
      require: function(moduleName) {
        if (moduleName === 'docx') {
          return { 
            Document, Packer, Paragraph, TextRun, 
            Table, TableRow, TableCell, 
            ImageRun,
            Header, Footer,
            HeadingLevel, AlignmentType, LevelFormat, BorderStyle, 
            PageOrientation,
            ExternalHyperlink, InternalHyperlink, Bookmark, 
            FootnoteReferenceRun,
            PositionalTab, PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader,
            TabStopType, TabStopPosition, 
            Column, SectionType,
            TableOfContents,
            WidthType, ShadingType,
            VerticalAlign,
            PageNumber, PageBreak 
          };
        }
        if (moduleName === 'fs') {
          return {
            writeFileSync: function(filePath, buffer) {
              // Redirect all writes to our output directory
              // Extract just the filename from whatever path the user provided
              const filename = path.basename(filePath);
              const safeOutputPath = path.join(OUTPUT_DIR, filename);
              
              if (NODE_ENV === 'development') {
                console.log(`Redirecting write from ${filePath} to ${safeOutputPath}`);
              }
              
              fs.writeFileSync(safeOutputPath, buffer);
              return buffer.length;  // Return bytes written
            }
          };
        }
        throw new Error(`Module '${moduleName}' is not allowed`);
      },
      console: {
        log: function(...args) { NODE_ENV === 'development' && console.log(...args); }
      },
      // Export all components for direct use
      Document, Packer, Paragraph, TextRun, 
      Table, TableRow, TableCell, 
      ImageRun,
      Header, Footer,
      HeadingLevel, AlignmentType, LevelFormat, BorderStyle, 
      PageOrientation,
      ExternalHyperlink, InternalHyperlink, Bookmark, 
      FootnoteReferenceRun,
      PositionalTab, PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader,
      TabStopType, TabStopPosition, 
      Column, SectionType,
      TableOfContents,
      WidthType, ShadingType,
      VerticalAlign,
      PageNumber, PageBreak 
    };

    // Validate code syntax before execution
    try {
      new vm.Script(jsCode);
    } catch (syntaxErr) {
      return res.status(400).json({ 
        error: `Code syntax error: ${syntaxErr.message}. Ensure your code is valid JavaScript.`
      });
    }

    // Execute code in sandbox with timeout - let user code create document as desired
    const context = vm.createContext(sandbox);
    const script = new vm.Script(jsCode);
    
    // Run the user's code to create the document
    try {
      script.runInContext(context, { timeout: 30000 });
    } catch (execError) {
      if (NODE_ENV === 'development') {
        console.error('Script execution error:', execError.message);
      }
      return res.status(500).json({
        error: `Script execution error: ${execError.message}`
      });
    }

    // Wait a bit for async operations (like Packer.toBuffer().then()) to complete
    await new Promise(resolve => setTimeout(resolve, 1000));

    // After execution, capture the document from sandbox
    // Try to find a Document instance that was created
    let doc = null;
    
    // Check common variable names first
    const commonNames = ['doc', 'document', 'myDoc', 'docx', 'result', 'output', 'mydoc', 'DOC', 'DOCUMENT', 'businessDoc', 'guideDoc', 'guide'];
    for (const varName of commonNames) {
      const val = context[varName];
      if (val && typeof val === 'object') {
        if (NODE_ENV === 'development') {
          console.log(`Checking '${varName}':`, val?.constructor?.name, 'has sections:', !!val.sections);
        }
        // Check if it's a Document (has sections property)
        if (val.sections !== undefined && Array.isArray(val.sections)) {
          doc = val;
          if (NODE_ENV === 'development') {
            console.log(`✓ Found Document in context.${varName}`);
          }
          break;
        }
      }
    }

    // If not found in common names, search all properties
    if (!doc) {
      if (NODE_ENV === 'development') {
        console.log('Document not found in common names. Searching all properties...');
      }
      for (const key in context) {
        if (key.startsWith('__') || typeof context[key] !== 'object') continue;
        
        const val = context[key];
        if (val && (val.sections !== undefined && Array.isArray(val.sections))) {
          doc = val;
          if (NODE_ENV === 'development') {
            console.log(`✓ Found Document in context.${key}`);
          }
          break;
        }
      }
    }

    // If still no document found, check if a file was written to output directory
    // This handles cases where the user code uses Packer.toBuffer().then() to write directly
    if (!doc) {
      if (NODE_ENV === 'development') {
        console.log('No Document object found. Checking if file was written to output directory...');
      }
      
      const files = fs.readdirSync(OUTPUT_DIR);
      const docxFiles = files.filter(f => f.endsWith('.docx'));
      
      if (docxFiles.length > 0) {
        // Use the most recently modified file
        const filePath = path.join(OUTPUT_DIR, docxFiles[0]);
        const stats = fs.statSync(filePath);
        
        if (NODE_ENV === 'development') {
          console.log(`✓ Found generated DOCX file: ${filePath} (${stats.size} bytes)`);
        }
        
        // Send the file
        res.download(filePath, 'generated.docx', (err) => {
          cleanupFiles(uploadPath, filePath);
          if (err && NODE_ENV === 'development') {
            console.error('Download error:', err);
          }
        });
        return;
      }
      
      if (NODE_ENV === 'development') {
        const allKeys = Object.keys(context).filter(k => typeof context[k] === 'object' && !k.startsWith('__'));
        console.log('No Document found. Object properties in context:', allKeys.slice(0, 30));
      }
      return res.status(400).json({ 
        error: 'No Document object found and no DOCX file was generated. Your code must create a Document instance or generate a valid DOCX file.'
      });
    }

    // Generate DOCX using Packer.toBuffer
    try {
      const buffer = await Packer.toBuffer(doc);
      
      if (!buffer || buffer.length === 0) {
        return res.status(500).json({ error: 'Generated DOCX buffer is empty' });
      }

      // Write to output path
      fs.writeFileSync(outputPath, buffer);
    } catch (packErr) {
      return res.status(500).json({ 
        error: `Failed to generate DOCX: ${packErr.message}`
      });
    }

    // Check if file was created
    if (!fs.existsSync(outputPath)) {
      return res.status(500).json({ error: 'DOCX file was not created. Please check your code.' });
    }

    const fileSize = fs.statSync(outputPath).size;
    if (fileSize === 0) {
      return res.status(500).json({ error: 'Generated DOCX file is empty' });
    }

    // Send the file
    res.download(outputPath, 'generated.docx', (err) => {
      cleanupFiles(uploadPath, outputPath);
      if (err && NODE_ENV === 'development') {
        console.error('Download error:', err);
      }
    });

  } catch (error) {
    console.error('Error:', error.message);
    cleanupFiles(uploadPath, outputPath);
    
    // Provide helpful error messages based on SKILL.md critical rules
    let errorMsg = error.message;
    
    // Common docx-js issues
    if (errorMsg.includes('Unexpected token')) {
      errorMsg = 'Invalid JavaScript syntax in uploaded file';
    } else if (errorMsg.includes('Document variable not found')) {
      errorMsg = 'Your code must create a Document instance. Example: const doc = new Document({...})';
    } else if (errorMsg.includes('is not a function')) {
      errorMsg = 'Make sure you\'re using: new Document({sections: [...]})';
    } else if (errorMsg.includes('ImageRun')) {
      errorMsg = 'ImageRun requires: type (png/jpg/jpeg/gif/bmp/svg), data, transformation, and altText parameters. See documentation.';
    } else if (errorMsg.includes('Table')) {
      errorMsg = 'Critical: Tables need BOTH columnWidths array AND cell width. Table width must equal sum of columnWidths (use WidthType.DXA, never PERCENTAGE). Use ShadingType.CLEAR for colors.';
    } else if (errorMsg.includes('PageBreak')) {
      errorMsg = 'PageBreak must be inside a Paragraph: new Paragraph({ children: [new PageBreak()] })';
    } else if (errorMsg.includes('Header') || errorMsg.includes('Footer')) {
      errorMsg = 'Headers/Footers go in sections.properties.headers/footers. See documentation for examples.';
    } else if (errorMsg.includes('TOC') || errorMsg.includes('TableOfContents')) {
      errorMsg = 'TableOfContents requires HeadingLevel on heading paragraphs with outlineLevel specified. See documentation.';
    } else if (errorMsg.includes('unicode') || errorMsg.includes('bullet')) {
      errorMsg = 'Never use unicode bullets (•). Use LevelFormat.BULLET with numbering config instead. See documentation.';
    } else if (errorMsg.includes('Page')) {
      errorMsg = 'CRITICAL: Set page size explicitly. Default is A4. For US Letter: width: 12240, height: 15840 (DXA units).';
    }
    
    res.status(500).json({ error: errorMsg || 'Internal server error' });
  }
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(err.status || 500).json({ error: err.message || 'Internal server error' });
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
  console.log(`📁 Upload directory: ${UPLOAD_DIR}`);
  console.log(`📁 Output directory: ${OUTPUT_DIR}`);
  console.log(`🌍 Environment: ${NODE_ENV}`);
});
