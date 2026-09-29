# AK Shop — Quick Setup

## 1. Install dependencies
```bash
npm install
```

## 2. Configure Supabase

1. Create a new Supabase project (or use existing)
2. Copy your project URL and anon key
3. Edit `.env`:
   ```
   EXPO_PUBLIC_SUPABASE_URL=https://your-project.supabase.co
   EXPO_PUBLIC_SUPABASE_ANON_KEY=your_anon_key_here
   ```

## 3. Run database migration

Open Supabase SQL Editor and run all SQL from **MIGRATION.md**.

This creates:
- `products` table
- RLS policies (public read, authenticated write)
- Indexes and triggers

## 4. Run the app

```bash
npm start
```

Then choose iOS/Android/Web.

## Project Structure

```
Shop/
├── app/
│   ├── _layout.tsx          # Root layout (expo-router)
│   └── (tabs)/
│       ├── _layout.tsx      # Tab navigator (Products, Add)
│       ├── index.tsx        # Products list (public view)
│       └── add.tsx          # Add product (shopkeeper)
├── lib/
│   ├── supabase.ts          # Supabase client
│   └── types.ts             # TypeScript types (Product, Shopkeeper)
├── .env                     # Supabase credentials (not committed)
├── MIGRATION.md             # Database SQL
└── README.md                # Full documentation
```

## Current Features

- ✅ View all products (public, no auth needed)
- ✅ Add products (requires Supabase auth — add login later)
- ✅ Pull to refresh
- ✅ Product cards with image, name, description, price, stock

## Next Steps

- Add authentication (login/signup for shopkeepers)
- Add product images (camera/gallery upload)
- Add search/filter by category
- Add product detail view
- Add edit/delete products
- Add shopkeeper profile/settings
