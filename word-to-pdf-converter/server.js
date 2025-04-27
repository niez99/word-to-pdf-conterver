const express = require('express');
const multer = require('multer');
const libre = require('libreoffice-convert');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));

function checkLibreOfficeInstalled(callback) {
  exec('soffice --version', (error, stdout, stderr) => {
    if (error) {
      callback(false);
    } else {
      callback(true);
    }
  });
}

app.post('/convert', upload.single('wordFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  checkLibreOfficeInstalled((isInstalled) => {
    if (!isInstalled) {
      return res.status(500).send('LibreOffice is not installed or not found in PATH.');
    }

    const wordFilePath = req.file.path;
    const ext = '.pdf';
    const inputFile = fs.readFileSync(wordFilePath);

    libre.convert(inputFile, ext, undefined, (err, done) => {
      if (err) {
        console.error(`Conversion error: ${err}`);
        return res.status(500).send('Failed to convert file.');
      }

      const pdfFileName = path.basename(req.file.originalname, path.extname(req.file.originalname)) + ext;

      res.set({
        'Content-Type': 'application/pdf',
        'Content-Disposition': 'attachment; filename="' + pdfFileName + '"',
      });
      res.send(done);

      fs.unlink(wordFilePath, (err) => {
        if (err) console.error('Failed to delete temp file:', err);
      });
    });
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
