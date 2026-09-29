# AK Shop — Product Capture & Publish Plan

## Vision (Steve Jobs–style)

One screen. One photo. One publish. The shopkeeper should feel like the product uploads itself.

---

## Goals

- Shopkeeper can take a product photo and publish an item quickly.
- AI can draft product details and shopkeeper can edit before publishing.
- Customers can see available items immediately.
- Settings for OpenAI/Gemini keys are simple and safe.

## Non-Goals (for this phase)

- Full authentication flow (login/signup UI)
- Payment or checkout
- Advanced inventory analytics

---

## System Plan (Step-by-step)

### Step 1 — Photo Capture

**What:** Add camera/gallery picker on Add Product screen.
**Why:** Photo first drives the rest of the flow.
**Status:** ✅ Done

### Step 2 — AI Auto‑fill

**What:** Use OpenAI or Gemini to draft name, description, price, category, stock.
**Why:** Save time; shopkeeper can edit.
**Status:** ✅ Done

### Step 3 — Publish Flow

**What:** Validate inputs, upload image to storage, write product to database.
**Why:** One tap to publish.
**Status:** ✅ Done

### Step 4 — Settings for AI

**What:** Store API keys and model choice on-device.
**Why:** Privacy and control.
**Status:** ✅ Done

### Step 5 — Public Catalog

**What:** Customers browse products without login.
**Why:** Frictionless viewing.
**Status:** ✅ Done (existing)

### Step 6 — Storage & Permissions

**What:** Create product-images bucket and policies in Supabase.
**Why:** Serve images publicly, upload securely.
**Status:** ⚠️ Needs SQL run in Supabase

---

## Project Manager Check (Gap Review)

### ✅ Ready

- Photo capture UI
- AI auto‑fill (OpenAI/Gemini)
- Publish flow with image upload
- Settings screen
- Public product listing

### ⚠️ Remaining

- Run storage SQL for bucket + policies
- Add shopkeeper authentication flow (optional but recommended)

### Risk

- AI auto‑fill depends on valid API keys.
- Image upload depends on storage bucket creation.

---

## Ship Checklist

- [ ] Run storage SQL (bucket + policies)
- [ ] Add AI keys in Settings
- [ ] Verify publish flow end‑to‑end

---

## Next Phase Ideas (Optional)

- Login/signup for shopkeepers
- Edit/delete products
- Product detail screen
- Search/filter by category

---

## Final Go/No‑Go

**Go** once storage SQL is run and an AI key is set. Everything else is complete.
