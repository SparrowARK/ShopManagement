import mongoose from "mongoose";
import { GridFSBucket, ObjectId } from "mongodb";

let bucket: GridFSBucket;

/**
 * Initialize the GridFS bucket after Mongoose connects.
 */
export function initGridFS(connection: mongoose.Connection): void {
  const db = connection.db;
  if (!db) {
    throw new Error("Database not available on connection");
  }
  bucket = new GridFSBucket(db, { bucketName: "images" });
  console.log("✅ GridFS bucket initialized");
}

/**
 * Get the GridFS bucket instance.
 */
export function getBucket(): GridFSBucket {
  if (!bucket) {
    throw new Error("GridFS bucket not initialized. Call initGridFS first.");
  }
  return bucket;
}

/**
 * Upload a file buffer to GridFS.
 * Returns the file ID as a string.
 */
export async function uploadToGridFS(
  buffer: Buffer,
  filename: string,
  contentType: string
): Promise<string> {
  return new Promise((resolve, reject) => {
    const uploadStream = getBucket().openUploadStream(filename, {
      contentType,
    });

    uploadStream.on("finish", () => {
      resolve(uploadStream.id.toString());
    });

    uploadStream.on("error", (err) => {
      reject(err);
    });

    uploadStream.end(buffer);
  });
}

/**
 * Delete a file from GridFS by its ID string.
 */
export async function deleteFromGridFS(fileId: string): Promise<void> {
  try {
    await getBucket().delete(new ObjectId(fileId));
  } catch (err) {
    // File may already be deleted — log but don't throw
    console.warn(`GridFS delete warning for ${fileId}:`, err);
  }
}
