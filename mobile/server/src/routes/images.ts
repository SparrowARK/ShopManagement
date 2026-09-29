import { Router, Request, Response } from "express";
import { ObjectId } from "mongodb";
import { getBucket } from "../utils/gridfs";

const router = Router();

/**
 * GET /api/images/:id
 * Public — serve an image from GridFS by its file ID.
 * Streams the binary directly to the response with caching headers.
 */
router.get("/:id", async (req: Request, res: Response): Promise<void> => {
  try {
    const fileId = new ObjectId(req.params.id as string);
    const bucket = getBucket();

    // Find the file metadata first
    const files = await bucket.find({ _id: fileId }).toArray();
    if (files.length === 0) {
      res.status(404).json({ error: "Image not found" });
      return;
    }

    const file = files[0];
    const contentType = file.contentType || "image/jpeg";

    // Set headers
    res.set("Content-Type", contentType);
    res.set("Cache-Control", "public, max-age=86400"); // Cache 24h
    res.set("Content-Length", file.length?.toString() || "");

    // Stream the file
    const downloadStream = bucket.openDownloadStream(fileId);

    downloadStream.on("error", (err) => {
      console.error("GridFS stream error:", err);
      if (!res.headersSent) {
        res.status(500).json({ error: "Failed to stream image" });
      }
    });

    downloadStream.pipe(res);
  } catch (err) {
    console.error("Image serve error:", err);
    if (!res.headersSent) {
      res.status(400).json({ error: "Invalid image ID" });
    }
  }
});

export default router;
