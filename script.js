const markdownInput = document.getElementById('markdownInput');
const preview = document.getElementById('preview');

if (window.marked) {
    window.marked.setOptions({ breaks: true, gfm: true });
}

markdownInput.addEventListener('input', updatePreview);

function escapeHtml(value) {
    return value.replace(/[&<>"']/g, (character) => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
    })[character]);
}

function updatePreview() {
    const markdown = markdownInput.value;
    if (!markdown.trim()) {
        preview.innerHTML = '';
        return;
    }

    if (window.marked && typeof window.marked.parse === 'function') {
        preview.innerHTML = window.marked.parse(markdown);
    } else {
        preview.innerHTML = `<pre>${escapeHtml(markdown)}</pre>`;
    }
}

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

*Generated with ❤️ by Zavoka Converter*`;
    updatePreview();
}

function clearAll() {
    markdownInput.value = '';
    preview.innerHTML = '';
}

function readMarkdownFile(file) {
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
        markdownInput.value = String(reader.result || '');
        updatePreview();
    };
    reader.onerror = () => alert('The file could not be read. Please try another file.');
    reader.readAsText(file);
}

function handleFile(input) {
    const file = input.files && input.files[0];
    if (file) readMarkdownFile(file);
    input.value = '';
}

async function convertToPDF() {
    if (!markdownInput.value.trim()) {
        alert('Please enter some markdown content first!');
        return;
    }
    if (typeof window.html2pdf !== 'function') {
        alert('PDF export is unavailable. Check that the PDF library loaded.');
        return;
    }

    try {
        const element = document.createElement('div');
        element.className = 'preview-content';
        element.innerHTML = window.marked && window.marked.parse
            ? window.marked.parse(markdownInput.value)
            : `<pre>${escapeHtml(markdownInput.value)}</pre>`;

        const options = {
            margin: 10,
            filename: 'document.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2 },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        };
        await window.html2pdf().set(options).from(element).save();
    } catch (error) {
        console.error('PDF export failed:', error);
        alert('PDF export failed. Please try again.');
    }
}

function markdownToDocxText(text) {
    return text
        .replace(/!\[([^\]]*)\]\([^)]*\)/g, '$1')
        .replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
        .replace(/\*\*|__|~~|`|\*|_/g, '')
        .trim();
}

async function convertToDOCX() {
    if (!markdownInput.value.trim()) {
        alert('Please enter some markdown content first!');
        return;
    }
    if (!window.docx || !window.docx.Document || !window.docx.Packer || !window.docx.Paragraph || !window.docx.HeadingLevel) {
        alert('DOCX export is unavailable because the DOCX library did not load. Check your internet connection and reload the page.');
        return;
    }

    try {
        const { Document, Packer, Paragraph, HeadingLevel } = window.docx;
        const children = markdownInput.value.split(/\r?\n/).map((line) => {
            const heading = line.match(/^(#{1,6})\s+(.*)$/);
            if (heading) {
                const level = Math.min(heading[1].length, 3);
                const headingLevel = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3][level - 1];
                return new Paragraph({ text: markdownToDocxText(heading[2]), heading: headingLevel });
            }

            const bullet = line.match(/^\s*[-*+]\s+(.*)$/);
            if (bullet) {
                return new Paragraph({ text: markdownToDocxText(bullet[1]), bullet: { level: 0 } });
            }

            const numbered = line.match(/^\s*(\d+)[.)]\s+(.*)$/);
            if (numbered) {
                return new Paragraph({ text: `${numbered[1]}. ${markdownToDocxText(numbered[2])}` });
            }

            if (/^\s*>\s?/.test(line)) {
                return new Paragraph({ text: markdownToDocxText(line.replace(/^\s*>\s?/, '')), style: 'IntenseQuote' });
            }

            if (/^\s*([-*_])\s*\1\s*\1+\s*$/.test(line)) {
                return new Paragraph({ text: '' });
            }

            return new Paragraph({ text: markdownToDocxText(line) });
        });

        const documentFile = new Document({ sections: [{ children }] });
        const blob = await Packer.toBlob(documentFile);
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'zavoka-document.docx';
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
        console.error('DOCX export failed:', error);
        alert('DOCX export failed. Please try again.');
    }
}

const container = document.querySelector('.container');
container.addEventListener('dragover', (event) => {
    event.preventDefault();
    container.style.outline = '3px dashed #667eea';
    container.style.outlineOffset = '-3px';
});

container.addEventListener('dragleave', (event) => {
    if (!container.contains(event.relatedTarget)) {
        container.style.outline = '';
    }
});

container.addEventListener('drop', (event) => {
    event.preventDefault();
    container.style.outline = '';

    const file = event.dataTransfer.files[0];
    if (file && /\.(md|txt)$/i.test(file.name)) {
        readMarkdownFile(file);
    } else if (file) {
        alert('Please drop a .md or .txt file!');
    }
});

loadSample();
