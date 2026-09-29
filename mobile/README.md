# AK Shop

Shopkeeper app: add products, customers view available products.

---

## Step-by-Step Setup Guide

Follow these steps in order to get AK Shop running.

### Step 1: Install Dependencies

Open terminal in the project folder (`C:\DriveD\Projects\Shop`) and run:

```bash
npm install
```

Wait for all packages to install. This may take 1-2 minutes.

---

### Step 2: Create Supabase Project

1. Go to [supabase.com](https://supabase.com) and sign in (or create an account)
2. Click **"New Project"**
3. Fill in:
   - **Name:** AK Shop (or any name you prefer)
   - **Database Password:** Create a strong password (save it somewhere safe)
   - **Region:** Choose closest to you
   - Click **"Create new project"**
4. Wait 1-2 minutes for the project to finish setting up

---

### Step 3: Get Supabase Credentials

1. In your Supabase project dashboard, click **Settings** (gear icon) → **API**
2. Find these two values:
   - **Project URL** (looks like `https://xxxxx.supabase.co`)
   - **anon public** key (long string starting with `eyJ...`)
3. Copy both values (you'll need them in the next step)

---

### Step 4: Configure Environment Variables

1. Open the file `.env` in the project folder (`C:\DriveD\Projects\Shop\.env`)
2. Paste your Supabase credentials:

```
EXPO_PUBLIC_SUPABASE_URL=https://your-project-id.supabase.co
EXPO_PUBLIC_SUPABASE_ANON_KEY=your_anon_key_here
```

Replace `your-project-id.supabase.co` with your actual Project URL, and `your_anon_key_here` with your actual anon key.

**Important:** Don't add quotes around the values, just paste them directly.

---

### Step 5: Create Database Table

1. In Supabase dashboard, click **SQL Editor** (left sidebar)
2. Click **"New query"**
3. Open the file `MIGRATION.md` in this project
4. Copy **all the SQL code** from `MIGRATION.md` (starts with `CREATE TABLE IF NOT EXISTS products`)
5. Paste it into the SQL Editor
6. Click **"Run"** (or press Ctrl+Enter)
7. You should see: **"Success. No rows returned"** — this means the table was created successfully

---

### Step 6: Test the App

1. In terminal, run:
   ```bash
   npm start
   ```
2. A QR code and menu will appear
3. Choose one:
   - **Press `a`** for Android (or scan QR with Expo Go app)
   - **Press `i`** for iOS (or scan QR with Camera app)
   - **Press `w`** for Web browser

---

### Step 7: Test Adding a Product (Optional)

**Note:** Adding products requires authentication. For now, you can:
- View products (Products tab) — works without login
- Add Product tab — will show an error until you add authentication

To test adding products, you'll need to add a login screen (next step).

---

## What's Working Now

✅ **Products tab:** View all products (public, no login needed)  
✅ **Add Product tab:** Form is ready (needs authentication to work)  
✅ **Pull to refresh:** Swipe down on products list to refresh  
✅ **Product cards:** Shows image, name, description, price, stock

---

## Next Steps (After Setup)

1. **Add authentication** — Login/signup screen for shopkeepers
2. **Add product images** — Camera/gallery upload
3. **Add search/filter** — Filter by category
4. **Product detail view** — Tap product to see full details
5. **Edit/delete products** — Shopkeeper can manage their products

---

## Troubleshooting

**"Supabase URL or Anon Key missing" warning:**
- Check `.env` file exists and has correct values
- Make sure no quotes around the values
- Restart the app (`npm start`)

**"Table products does not exist" error:**
- Go back to Step 5 and run the SQL migration again
- Check Supabase SQL Editor → Tables → you should see `products` table

**"Not logged in" when adding product:**
- This is expected — add authentication screen first (see Next Steps)

**App won't start:**
- Make sure you ran `npm install` (Step 1)
- Check terminal for error messages
- Try deleting `node_modules` folder and running `npm install` again

---

## Project Structure

```
Shop/
├── app/
│   ├── _layout.tsx          # Root layout
│   └── (tabs)/
│       ├── _layout.tsx      # Tab navigator
│       ├── index.tsx        # Products list
│       └── add.tsx          # Add product form
├── lib/
│   ├── supabase.ts          # Supabase client
│   └── types.ts             # TypeScript types
├── .env                     # Your Supabase keys (not committed)
├── MIGRATION.md             # Database SQL
└── README.md                # This file
```

---

## Quick Reference

**Start app:**
```bash
npm start
```

**Check TypeScript:**
```bash
npx tsc --noEmit
```

**View Supabase tables:**
- Supabase Dashboard → Table Editor → `products`

**View Supabase SQL:**
- Supabase Dashboard → SQL Editor

---

*Ready to start? Begin with Step 1 above!*
