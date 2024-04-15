import express, { Application, Request, Response } from "express";

import cors from "cors";
import dotenv from "dotenv";
import multer from "multer";
import fs from "fs/promises";
import path from 'path';


dotenv.config();

const app: Application = express();
app.use(cors());

const port: number = parseInt(process.env.PORT || "3001");

app.get("/", (req: Request, res: Response) => {
  res.send("Grader service is running.");
});


const uploadsDir = path.join(__dirname, 'uploads');

async function createUploadsDirectory() {
    try {
        await fs.mkdir(uploadsDir, { recursive: true });
        console.log('Uploads directory created');
    } catch (err) {
        console.error('Failed to create directory:', err);
    }
}

createUploadsDirectory();
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, "uploads/"),
  filename: (req, file, cb) => {
    console.log("file1:", file);
    console.log("file name:", file.originalname);
  
    cb(null, `${file.originalname}`);
  }
});

const upload = multer({ storage });

app.post("/", upload.single("file"), async (req: Request, res: Response) => {
  if (!req.file) {
    return res.status(400).send("No file uploaded.");
  }
  const filePath = req.file.path; // Path to the stored file

  console.log("filePath:", filePath)
  try {
    const data = await fs.readFile(filePath, "utf8");
    console.log("data:", data)
    res.setHeader("Content-Type", "text/xml");
    res.send(data);
  } catch (error) {
    console.error("Error reading the file:", error);
    res.status(500).send("Failed to process the file");
  } finally {
    // Always clean up the uploaded file
    try {
      await fs.unlink(filePath);
    } catch (cleanupError) {
      console.error("Failed to clean up the file:", cleanupError);
    }
  }
});

app.listen(port, function () {
  console.log(`App is listening on port ${port} !`);
});
