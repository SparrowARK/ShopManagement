# Database Migration — AK Shop

Run these SQL statements in your Supabase SQL Editor.

## 1. Products Table

```sql
-- Products table
CREATE TABLE IF NOT EXISTS products (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  shopkeeper_id UUID REFERENCES auth.users(id) ON DELETE CASCADE NOT NULL,
  name TEXT NOT NULL,
  description TEXT,
  price NUMERIC(10, 2) NOT NULL CHECK (price >= 0),
  stock INTEGER CHECK (stock >= 0),
  category TEXT,
  image_url TEXT,
  created_at TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  updated_at TIMESTAMPTZ DEFAULT NOW() NOT NULL
);

-- Enable RLS
ALTER TABLE products ENABLE ROW LEVEL SECURITY;

-- Policy: Anyone can view products (public catalog)
CREATE POLICY "Anyone can view products" ON products
  FOR SELECT USING (true);

-- Policy: Only authenticated users can add products (shopkeepers)
CREATE POLICY "Shopkeepers can add products" ON products
  FOR INSERT WITH CHECK (auth.uid() = shopkeeper_id);

-- Policy: Only shopkeeper can update their own products
CREATE POLICY "Shopkeepers can update own products" ON products
  FOR UPDATE USING (auth.uid() = shopkeeper_id);

-- Policy: Only shopkeeper can delete their own products
CREATE POLICY "Shopkeepers can delete own products" ON products
  FOR DELETE USING (auth.uid() = shopkeeper_id);

-- Indexes for performance
CREATE INDEX IF NOT EXISTS idx_products_shopkeeper ON products(shopkeeper_id);
CREATE INDEX IF NOT EXISTS idx_products_created_at ON products(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_products_category ON products(category) WHERE category IS NOT NULL;

-- Trigger to update updated_at timestamp
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER update_products_updated_at
  BEFORE UPDATE ON products
  FOR EACH ROW
  EXECUTE FUNCTION update_updated_at_column();
```

## Notes

- **Public read:** Anyone can view products (no auth required for browsing).
- **Authenticated write:** Only logged-in users (shopkeepers) can add/edit/delete products.
- **Ownership:** Each product is tied to `shopkeeper_id` (the user who created it).
- **Price:** Stored as NUMERIC(10, 2) for precision (e.g., ₹999.99).
- **Stock:** Optional integer; null means "unlimited" or "not tracked".

## 2. Storage Bucket (Product Images)

Create a public bucket for product photos and allow public read, authenticated upload.

```sql
-- Create bucket
INSERT INTO storage.buckets (id, name, public)
VALUES ('product-images', 'product-images', true)
ON CONFLICT (id) DO NOTHING;

-- Public read access
CREATE POLICY "Public can view product images" ON storage.objects
  FOR SELECT USING (bucket_id = 'product-images');

-- Authenticated users can upload their own images
CREATE POLICY "Authenticated users can upload product images" ON storage.objects
  FOR INSERT WITH CHECK (
    bucket_id = 'product-images'
    AND auth.role() = 'authenticated'
  );

-- Authenticated users can delete their own images
CREATE POLICY "Authenticated users can delete product images" ON storage.objects
  FOR DELETE USING (
    bucket_id = 'product-images'
    AND auth.role() = 'authenticated'
  );
```
