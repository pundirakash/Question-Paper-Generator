const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const express = require('express');

async function generateDocxFromTemplate(data) {
  try {
    const templateFilePath = path.resolve(__dirname, 'template.docx');
    const outputDocxFilePath = path.resolve(__dirname, 'output.docx');

    const content = fs.readFileSync(templateFilePath, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    doc.setData(data);
    doc.render();

    const buffer = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(outputDocxFilePath, buffer);

    return outputDocxFilePath;
  } catch (error) {
    throw new Error(`Error generating DOCX file: ${error.message}`);
  }
}

// Example usage with express.js
const app = express();
app.use(express.json());

app.post('/generate-docx', async (req, res) => {
  try {
    const data = req.body;
    const outputDocxFilePath = await generateDocxFromTemplate(data);

    res.download(outputDocxFilePath, (err) => {
      if (err) {
        console.error(`Error sending the file: ${err.message}`);
      }
      // Optionally, delete the file after sending it
      // fs.unlinkSync(outputDocxFilePath);
    });
  } catch (error) {
    res.status(500).json({ message: `Server error: ${error.message}` });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
