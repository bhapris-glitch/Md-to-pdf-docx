// Initialize marked
marked.setOptions({
    breaks: true,
    gfm: true
});

// DOM elements
const markdownInput = document.getElementById('markdownInput');
const preview = document.getElementById('preview');

// Live preview
markdownInput.addEventListener('input', updatePreview);

function updatePreview() {
    const markdown = markdownInput.value;
    preview.innerHTML = marked.parse(markdown);
}

// Sample content
function loadSample() {
    markdownInput.value = `# 📄 Sample Document

## Introduction
This is a **sample** markdown document that demonstrates the converter.

### Features:
- Convert to **PDF**
- Convert to **DOCX**
- Live preview
- Drag & drop support

## Code Example
\`\`\`javascript
function hello() {
    console.log("Hello, World!");
}
\`\`\`

## Table Example
| Feature | Status |
|---------|--------|
| PDF Export | ✅ |
| DOCX Export | ✅ |
| Live Preview | ✅ |

> **Tip:** You can upload your own .md files or start typing!

---

*Generated with ❤️ by Document Converter*`;
    updatePreview();
}

// Clear all
function clearAll() {
    markdownInput.value = '';
    preview.innerHTML = '';
}

// File upload
function handleFile(input) {
    const file = input.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
        markdownInput.value = e.target.result;
        updatePreview();
    };
    reader.readAsText(file);
    input.value = '';
}

// Convert to PDF
function convertToPDF() {
    if (!markdownInput.value.trim()) {
        alert('Please enter some markdown content first!');
        return;
    }
    
    const element = document.createElement('div');
    element.innerHTML = marked.parse(markdownInput.value);
    element.style.padding = '20px';
    
    const opt = {
        margin: 10,
        filename: 'document.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };
    
    html2pdf().set(opt).from(element).save();
}

// Convert to DOCX
function convertToDOCX() {
    if (!markdownInput.value.trim()) {
        alert('Please enter some markdown content first!');
        return;
    }
    
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType } = docx;
    
    // Parse markdown to simple structure
    const lines = markdownInput.value.split('\n');
    const children = [];
    
    for (const line of lines) {
        if (line.startsWith('# ')) {
            children.push(new Paragraph({
                text: line.substring(2),
                heading: HeadingLevel.HEADING_1
            }));
        } else if (line.startsWith('## ')) {
            children.push(new Paragraph({
                text: line.substring(3),
                heading: HeadingLevel.HEADING_2
            }));
        } else if (line.startsWith('### ')) {
            children.push(new Paragraph({
                text: line.substring(4),
                heading: HeadingLevel.HEADING_3
            }));
        } else if (line.startsWith('- ')) {
            children.push(new Paragraph({
                text: line.substring(2),
                bullet: { level: 0 }
            }));
        } else if (line.startsWith('> ')) {
            children.push(new Paragraph({
                text: line.substring(2),
                style: 'IntenseQuote'
            }));
        } else if (line.trim() === '') {
            children.push(new Paragraph({ text: '' }));
        } else {
            children.push(new Paragraph({
                children: [new TextRun(line)]
            }));
        }
    }
    
    const doc = new Document({
        sections: [{
            properties: {},
            children: children
        }]
    });
    
    Packer.toBlob(doc).then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'document.docx';
        a.click();
        window.URL.revokeObjectURL(url);
    });
}

// Initialize with sample
loadSample();

// Drag and drop support
const container = document.querySelector('.container');
container.addEventListener('dragover', (e) => {
    e.preventDefault();
    container.style.border = '3px dashed #667eea';
});

container.addEventListener('dragleave', () => {
    container.style.border = 'none';
});

container.addEventListener('drop', (e) => {
    e.preventDefault();
    container.style.border = 'none';
    
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.md') || file.name.endsWith('.txt'))) {
        const reader = new FileReader();
        reader.onload = (event) => {
            markdownInput.value = event.target.result;
            updatePreview();
        };
        reader.readAsText(file);
    } else {
        alert('Please drop a .md or .txt file!');
    }
});
