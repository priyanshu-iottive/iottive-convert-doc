import type { Express } from "express";
import type { Server } from "http";
import multer from "multer";
import path from "path";
import fs from "fs";
import { parseDocx, parsePdf, buildBrandedDocx } from "./converter";

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if ([".pdf", ".docx"].includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error("Only PDF and DOCX files are supported"));
    }
  },
});

export async function registerRoutes(server: Server, app: Express) {
  // Convert uploaded document to branded DOCX
  app.post("/api/convert", upload.single("file"), async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: "No file uploaded" });
      }

      const clientName = (req.body.clientName as string) || "Client Name";
      const contactName = (req.body.contactName as string) || "Contact Person";
      const contactEmail = (req.body.contactEmail as string) || "email@example.com";
      const projectTitle = (req.body.projectTitle as string) || "Technical & Commercial Proposal";

      const ext = path.extname(req.file.originalname).toLowerCase();
      let parsed;

      if (ext === ".docx") {
        parsed = await parseDocx(req.file.buffer);
      } else if (ext === ".pdf") {
        parsed = await parsePdf(req.file.buffer);
      } else {
        return res.status(400).json({ error: "Unsupported file format" });
      }

      const brandedDocx = await buildBrandedDocx(
        parsed,
        clientName,
        contactName,
        contactEmail,
        projectTitle
      );

      const outputName = `IOTTIVE_${clientName.replace(/\s+/g, "_")}_${projectTitle.replace(/\s+/g, "_")}.docx`;

      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename="${outputName}"`);
      res.send(brandedDocx);
    } catch (error: any) {
      console.error("Conversion error:", error);
      res.status(500).json({ error: error.message || "Conversion failed" });
    }
  });

  // Health check
  app.get("/api/health", (_req, res) => {
    res.json({ status: "ok" });
  });
}
