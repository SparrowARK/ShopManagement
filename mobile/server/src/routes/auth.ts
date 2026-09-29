import { Router, Request, Response } from "express";
import { User } from "../models/User";
import { signToken, requireAuth } from "../middleware/auth";

const router = Router();

/**
 * POST /api/auth/register
 * Create a new account (shopkeeper or customer).
 */
router.post("/register", async (req: Request, res: Response): Promise<void> => {
  try {
    const { email, password, name, role, shop_name } = req.body;

    // Validate required fields
    if (!email || !password || !name || !role) {
      res.status(400).json({ error: "email, password, name, and role are required" });
      return;
    }

    if (!["shopkeeper", "customer"].includes(role)) {
      res.status(400).json({ error: "role must be 'shopkeeper' or 'customer'" });
      return;
    }

    if (password.length < 6) {
      res.status(400).json({ error: "Password must be at least 6 characters" });
      return;
    }

    // Check if email already exists
    const existing = await User.findOne({ email: email.toLowerCase().trim() });
    if (existing) {
      res.status(409).json({ error: "Email already registered" });
      return;
    }

    // Create user (password is hashed by pre-save hook)
    const user = await User.create({
      email: email.toLowerCase().trim(),
      password,
      name: name.trim(),
      role,
      shop_name: role === "shopkeeper" ? (shop_name?.trim() || null) : null,
    });

    const token = signToken(user);

    res.status(201).json({
      token,
      user: user.toJSON(),
    });
  } catch (err) {
    console.error("Register error:", err);
    res.status(500).json({ error: "Failed to register" });
  }
});

/**
 * POST /api/auth/login
 * Login with email + password, receive JWT.
 */
router.post("/login", async (req: Request, res: Response): Promise<void> => {
  try {
    const { email, password } = req.body;

    if (!email || !password) {
      res.status(400).json({ error: "email and password are required" });
      return;
    }

    const user = await User.findOne({ email: email.toLowerCase().trim() });
    if (!user) {
      res.status(401).json({ error: "Invalid email or password" });
      return;
    }

    const isMatch = await user.comparePassword(password);
    if (!isMatch) {
      res.status(401).json({ error: "Invalid email or password" });
      return;
    }

    const token = signToken(user);

    res.json({
      token,
      user: user.toJSON(),
    });
  } catch (err) {
    console.error("Login error:", err);
    res.status(500).json({ error: "Failed to login" });
  }
});

/**
 * GET /api/auth/me
 * Get current user from JWT token.
 */
router.get("/me", requireAuth, async (req: Request, res: Response): Promise<void> => {
  res.json({ user: req.user!.toJSON() });
});

export default router;
