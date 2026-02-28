import express, { Request, Response, NextFunction } from "express";
import multer from "multer";
import XLSX from "xlsx";
import fs from "fs/promises";
import path from "path";

const app = express();
const PORT = 3000;

const uploadDir = path.join(__dirname, "../uploads");

// Ensure uploads directory exists
async function ensureUploadDir(): Promise<void> {
  try {
    await fs.mkdir(uploadDir, { recursive: true });
  } catch (error) {
    console.error("Failed to create uploads directory:", error);
    process.exit(1);
  }
}

// Multer memory storage
const upload = multer({
  storage: multer.memoryStorage(),
  fileFilter: (req, file, cb) => {
    if (file.originalname.includes(".xls")) {
      cb(null, true);
    } else {
      cb(new Error("Only .xlsx files are allowed"));
    }
  },
});

// POST endpoint
app.post(
  "/upload",
  upload.single("file"),
  async (req: Request, res: Response, next: NextFunction) => {
    try {
      if (!req.file) {
        return res.status(400).send("No file uploaded");
      }

      // Load workbook from memory
      const workbook = XLSX.read(req.file.buffer, { type: "buffer" });

      // Generate new filename
      const newFileName = `copy_${Date.now()}.xlsx`;
      const newFilePath = path.join(uploadDir, newFileName);

      // Save copied file
      XLSX.writeFile(workbook, newFilePath);

      // Send copied file
      res.download(newFilePath, newFileName, async (err) => {
        if (err) {
          console.error("Download error:", err);
        }

        // Cleanup file after sending
        try {
          await fs.unlink(newFilePath);
        } catch (cleanupError) {
          console.error("Cleanup error:", cleanupError);
        }
      });
    } catch (error) {
      next(error);
    }
  }
);

// Global error handler
app.use((err: Error, req: Request, res: Response, next: NextFunction) => {
  console.error(err);
  res.status(500).send(err.message || "Internal Server Error");
});

ensureUploadDir().then(() => {
  app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
  });
});
