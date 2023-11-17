const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const path = require('path');
const cors=require('cors')
const app = express();
app.use(cors())
const PORT = process.env.PORT || 5000;

// Connect to MongoDB
mongoose.connect('mongodb://localhost:27017/mern_file_upload', {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

// Create a mongoose model for the file schema
const File = mongoose.model('File', {
  fileName: String,
  filePath: String,
});

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.resolve(__dirname, 'uploads/'));
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage: storage });


// Route for file upload
app.post('/upload', upload.single('file'), async (req, res) => {
    try {
      console.log('yash',req.file);
    const { originalname: fileName, path: filePath } = req.file;
    const { name } = req.body;

    const file = new File({
      fileName: fileName || name,
      filePath,
    });

    await file.save();
    res.status(201).send('File Uploaded Successfully');
  } catch (error) {
    res.status(500).send(error.message);
  }
});

// Route for file download
app.get('/download/:id', async (req, res) => {
  try {
    const file = await File.findById(req.params.id);
    res.download(file.filePath);
  } catch (error) {
    res.status(500).send(error.message);
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
