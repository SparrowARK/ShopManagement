import { Router, Request, Response } from "express";
import multer from "multer";
import { Product } from "../models/Product";
import { requireAuth, requireRole } from "../middleware/auth";
import { uploadToGridFS, deleteFromGridFS } from "../utils/gridfs";

const router = Router();

// Multer: store in memory, max 5MB
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    if (file.mimetype.startsWith("image/")) {
      cb(null, true);
    } else {
      cb(new Error("Only image files are allowed"));
    }
  },
});

/**
 * GET /api/products
 * Public — list all products, newest first.
 * Optional query: ?category=Electronics&search=phone
 */
router.get("/", async (req: Request, res: Response): Promise<void> => {
  try {
    const filter: Record<string, unknown> = {};

    if (req.query.category) {
      filter.category = req.query.category;
    }
    if (req.query.search) {
      filter.name = { $regex: req.query.search, $options: "i" };
    }
    if (req.query.shopkeeper_id) {
      filter.shopkeeper_id = req.query.shopkeeper_id;
    }

    const products = await Product.find(filter)
      .sort({ created_at: -1 })
      .lean();

    res.json({ products });
  } catch (err) {
    console.error("Fetch products error:", err);
    res.status(500).json({ error: "Failed to fetch products" });
  }
});

/**
 * GET /api/products/:id
 * Public — get a single product.
 */
router.get("/:id", async (req: Request, res: Response): Promise<void> => {
  try {
    const product = await Product.findById(req.params.id).lean();
    if (!product) {
      res.status(404).json({ error: "Product not found" });
      return;
    }
    res.json({ product });
  } catch (err) {
    console.error("Fetch product error:", err);
    res.status(500).json({ error: "Failed to fetch product" });
  }
});

/**
 * POST /api/products
 * Shopkeeper only — create a product with optional image.
 * Accepts multipart/form-data (fields + image file).
 */
router.post(
  "/",
  requireAuth,
  requireRole("shopkeeper"),
  upload.single("image"),
  async (req: Request, res: Response): Promise<void> => {
    try {
      const { name, description, price, stock, category, image_url } = req.body;

      if (!name?.trim() || price === undefined || price === "") {
        res.status(400).json({ error: "name and price are required" });
        return;
      }

      const priceNum = parseFloat(price);
      if (isNaN(priceNum) || priceNum < 0) {
        res.status(400).json({ error: "price must be a valid non-negative number" });
        return;
      }

      let imageId: string | null = null;

      // If an image file was uploaded, store in GridFS
      if (req.file) {
        imageId = await uploadToGridFS(
          req.file.buffer,
          `${Date.now()}-${req.file.originalname}`,
          req.file.mimetype
        );
      }

      const stockParsed = parseFloat(stock);
      const stockNum =
        stock !== undefined && stock !== "" && Number.isFinite(stockParsed)
          ? Math.floor(Math.abs(stockParsed))
          : null;

      const product = await Product.create({
        shopkeeper_id: req.user!._id,
        name: name.trim(),
        description: description?.trim() || null,
        price: priceNum,
        stock: stockNum,
        category: category?.trim() || null,
        image_id: imageId || image_url?.trim() || null,
      });

      res.status(201).json({ product: product.toJSON() });
    } catch (err) {
      console.error("Create product error:", err);
      res.status(500).json({ error: "Failed to create product" });
    }
  }
);

/**
 * PUT /api/products/:id
 * Owner only — update a product.
 */
router.put(
  "/:id",
  requireAuth,
  requireRole("shopkeeper"),
  upload.single("image"),
  async (req: Request, res: Response): Promise<void> => {
    try {
      const product = await Product.findById(req.params.id);
      if (!product) {
        res.status(404).json({ error: "Product not found" });
        return;
      }

      // Only the owner can update
      if (product.shopkeeper_id.toString() !== req.user!._id.toString()) {
        res.status(403).json({ error: "Not your product" });
        return;
      }

      const updates: Record<string, unknown> = {};

      if (req.body.name !== undefined) updates.name = req.body.name.trim();
      if (req.body.description !== undefined)
        updates.description = req.body.description.trim() || null;
      if (req.body.price !== undefined) {
        const p = parseFloat(req.body.price);
        if (!isNaN(p) && p >= 0) updates.price = p;
      }
      if (req.body.stock !== undefined) {
        const s = parseFloat(req.body.stock);
        updates.stock =
          req.body.stock !== "" && Number.isFinite(s)
            ? Math.floor(Math.abs(s))
            : null;
      }
      if (req.body.category !== undefined)
        updates.category = req.body.category.trim() || null;

      // Handle new image upload
      if (req.file) {
        // Delete old image from GridFS if it exists
        if (product.image_id) {
          await deleteFromGridFS(product.image_id);
        }
        updates.image_id = await uploadToGridFS(
          req.file.buffer,
          `${Date.now()}-${req.file.originalname}`,
          req.file.mimetype
        );
      }

      const updated = await Product.findByIdAndUpdate(
        req.params.id,
        { $set: updates },
        { new: true }
      ).lean();

      res.json({ product: updated });
    } catch (err) {
      console.error("Update product error:", err);
      res.status(500).json({ error: "Failed to update product" });
    }
  }
);

/**
 * DELETE /api/products/:id
 * Owner only — delete a product and its image.
 */
router.delete(
  "/:id",
  requireAuth,
  requireRole("shopkeeper"),
  async (req: Request, res: Response): Promise<void> => {
    try {
      const product = await Product.findById(req.params.id);
      if (!product) {
        res.status(404).json({ error: "Product not found" });
        return;
      }

      if (product.shopkeeper_id.toString() !== req.user!._id.toString()) {
        res.status(403).json({ error: "Not your product" });
        return;
      }

      // Delete image from GridFS
      if (product.image_id) {
        await deleteFromGridFS(product.image_id);
      }

      await Product.findByIdAndDelete(req.params.id);

      res.json({ message: "Product deleted" });
    } catch (err) {
      console.error("Delete product error:", err);
      res.status(500).json({ error: "Failed to delete product" });
    }
  }
);

export default router;
