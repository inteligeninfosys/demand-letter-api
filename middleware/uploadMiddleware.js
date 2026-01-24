import multer from 'multer';
import path from 'path';

const storage = multer.memoryStorage(); // we’ll stream to MinIO directly

export const upload = multer({
  storage,
  limits: {
    fileSize: 25 * 1024 * 1024, // 25MB
  },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const allowed = ['.pdf', '.docx', '.zip'];

    if (!allowed.includes(ext)) {
      return cb(new Error('Only .pdf, .docx, .zip files are allowed'));
    }

    cb(null, true);
  },
});
