import express, { Application, Request, Response } from "express";

import cors from "cors";
import dotenv from "dotenv";
import multer from "multer";
// import fs from "fs/promises";
import { Word2013CMLGenerator } from "./WordReader2013/Word2013CMLGenerator";
import { Word2013Translator } from "./WordReader2013/Word2013Translator";
import xmldom from "xmldom";

dotenv.config();

const app: Application = express();
app.use(cors());

const port: number = parseInt(process.env.PORT || "3001");

app.get("/", (req: Request, res: Response) => {
  res.send("Grader service is running.");
});

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, "uploads/"),
  filename: (req, file, cb) => cb(null, `${file.originalname}`),
});

const upload = multer({ storage });

app.post("/", upload.single("file"), async (req: Request, res: Response) => {
  if (!req.file) {
    return res.status(400).send("No file uploaded.");
  }
  try {
    const filePath = req.file.path; // Path to the stored file
    const cmlTranslator = new Word2013Translator();
    const cmlGenerator = new Word2013CMLGenerator();
    const docName = req.file.filename;

    const data = await cmlGenerator.generateCML(
      cmlTranslator,
      filePath,
      docName
    );
    // Create an XML Serializer to convert the document to a string
    const serializer = new xmldom.XMLSerializer();

    // Serialize the document to get the XML string
    const xmlString = serializer.serializeToString(data);
    res.setHeader("Content-Type", "text/xml");
    res.send(xmlString);
  } catch (error) {
    console.error("Error reading the file:", error);
    res.status(500).send(error || "Failed to process the file");
  } finally {
    // Always clean up the uploaded file
    // try {
    //   await fs.unlink(filePath);
    // } catch (cleanupError) {
    //   console.error("Failed to clean up the file:", cleanupError);
    // }
  }
});

app.listen(port, function () {
  console.log(`App is listening on port ${port} !`);
});
