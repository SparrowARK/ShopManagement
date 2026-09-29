import { Router, Request, Response } from "express";
import { Order } from "../models/Order";
import { Product } from "../models/Product";
import { requireAuth, requireRole } from "../middleware/auth";

const router = Router();

/**
 * POST /api/orders
 * Customer only — place an order.
 * Body: { items: [{ product_id, quantity }], customer_note? }
 */
router.post(
  "/",
  requireAuth,
  requireRole("customer"),
  async (req: Request, res: Response): Promise<void> => {
    try {
      const { items, customer_note } = req.body;

      if (!items || !Array.isArray(items) || items.length === 0) {
        res.status(400).json({ error: "items array is required and must not be empty" });
        return;
      }

      // Validate and snapshot each item
      const orderItems = [];
      let total = 0;
      let shopkeeperId: string | null = null;

      for (const item of items) {
        if (!item.product_id || !item.quantity || item.quantity < 1) {
          res.status(400).json({ error: "Each item needs product_id and quantity >= 1" });
          return;
        }

        const product = await Product.findById(item.product_id);
        if (!product) {
          res.status(404).json({ error: `Product ${item.product_id} not found` });
          return;
        }

        // All items in one order must belong to the same shopkeeper
        if (shopkeeperId === null) {
          shopkeeperId = product.shopkeeper_id.toString();
        } else if (product.shopkeeper_id.toString() !== shopkeeperId) {
          res.status(400).json({
            error: "All items in an order must be from the same shop",
          });
          return;
        }

        // Check stock if tracked
        if (product.stock !== null && product.stock !== undefined) {
          if (product.stock < item.quantity) {
            res.status(400).json({
              error: `Not enough stock for "${product.name}". Available: ${product.stock}`,
            });
            return;
          }
        }

        const lineTotal = product.price * item.quantity;
        total += lineTotal;

        orderItems.push({
          product_id: product._id,
          name: product.name,
          price: product.price,
          quantity: item.quantity,
        });
      }

      // Create the order
      const order = await Order.create({
        customer_id: req.user!._id,
        shopkeeper_id: shopkeeperId,
        items: orderItems,
        total,
        customer_note: customer_note?.trim() || null,
      });

      // Decrement stock for products that track it
      for (const item of orderItems) {
        await Product.findByIdAndUpdate(item.product_id, {
          $inc: { stock: -item.quantity },
        });
      }

      res.status(201).json({ order: order.toJSON() });
    } catch (err) {
      console.error("Create order error:", err);
      res.status(500).json({ error: "Failed to place order" });
    }
  }
);

/**
 * GET /api/orders
 * Authenticated — returns orders based on role.
 * Customer sees their orders. Shopkeeper sees orders for their products.
 */
router.get("/", requireAuth, async (req: Request, res: Response): Promise<void> => {
  try {
    const filter: Record<string, unknown> = {};

    if (req.user!.role === "customer") {
      filter.customer_id = req.user!._id;
    } else if (req.user!.role === "shopkeeper") {
      filter.shopkeeper_id = req.user!._id;
    }

    if (req.query.status) {
      filter.status = req.query.status;
    }

    const orders = await Order.find(filter)
      .sort({ created_at: -1 })
      .populate("customer_id", "name email")
      .lean();

    res.json({ orders });
  } catch (err) {
    console.error("Fetch orders error:", err);
    res.status(500).json({ error: "Failed to fetch orders" });
  }
});

/**
 * GET /api/orders/:id
 * Authenticated — get a single order (must be owner or shopkeeper).
 */
router.get("/:id", requireAuth, async (req: Request, res: Response): Promise<void> => {
  try {
    const order = await Order.findById(req.params.id)
      .populate("customer_id", "name email")
      .lean();

    if (!order) {
      res.status(404).json({ error: "Order not found" });
      return;
    }

    // Only the customer or shopkeeper can view
    const userId = req.user!._id.toString();
    const isOwner = order.customer_id &&
      typeof order.customer_id === "object" &&
      "_id" in order.customer_id &&
      (order.customer_id as { _id: { toString(): string } })._id.toString() === userId;
    const isShopkeeper = order.shopkeeper_id.toString() === userId;

    if (!isOwner && !isShopkeeper) {
      res.status(403).json({ error: "Not authorized" });
      return;
    }

    res.json({ order });
  } catch (err) {
    console.error("Fetch order error:", err);
    res.status(500).json({ error: "Failed to fetch order" });
  }
});

/**
 * PATCH /api/orders/:id/status
 * Shopkeeper only — update order status.
 * Body: { status: "confirmed" | "ready" | "delivered" | "cancelled" }
 */
router.patch(
  "/:id/status",
  requireAuth,
  requireRole("shopkeeper"),
  async (req: Request, res: Response): Promise<void> => {
    try {
      const { status } = req.body;
      const validStatuses = ["pending", "confirmed", "ready", "delivered", "cancelled"];

      if (!status || !validStatuses.includes(status)) {
        res.status(400).json({ error: `status must be one of: ${validStatuses.join(", ")}` });
        return;
      }

      const order = await Order.findById(req.params.id);
      if (!order) {
        res.status(404).json({ error: "Order not found" });
        return;
      }

      // Only the shopkeeper whose products were ordered can update
      if (order.shopkeeper_id.toString() !== req.user!._id.toString()) {
        res.status(403).json({ error: "Not your order to manage" });
        return;
      }

      // If cancelling, restore stock
      if (status === "cancelled" && order.status !== "cancelled") {
        for (const item of order.items) {
          await Product.findByIdAndUpdate(item.product_id, {
            $inc: { stock: item.quantity },
          });
        }
      }

      const updated = await Order.findByIdAndUpdate(
        req.params.id,
        { $set: { status, updated_at: new Date() } },
        { new: true }
      )
        .populate("customer_id", "name email")
        .lean();

      res.json({ order: updated });
    } catch (err) {
      console.error("Update order status error:", err);
      res.status(500).json({ error: "Failed to update order status" });
    }
  }
);

export default router;
