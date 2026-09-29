# AK Shop API on Vsys (Vercel)

Shop backend routes live in the Vsys Next.js app under `/api/ak-shop/*`.
Same MongoDB Atlas DB. No separate Express host needed.

## 1. Vercel env vars (VSys project)

Add these in Vercel → Project → Settings → Environment Variables:

| Name | Value |
|------|--------|
| `AK_SHOP_MONGODB_URI` | same as Shop `server/.env` → `MONGODB_URI` |
| `AK_SHOP_JWT_SECRET` | same as Shop `server/.env` → `JWT_SECRET` |

Redeploy after saving.

## 2. Atlas Network Access

Allow `0.0.0.0/0` (or Vercel’s IPs) so serverless functions can reach MongoDB.

## 3. Mobile app `.env`

```
EXPO_PUBLIC_API_URL=https://www.arkarki.com.np/api/ak-shop
```

Restart Expo (`npx expo start -c`).

## 4. Smoke test

```
GET https://www.arkarki.com.np/api/ak-shop/health
```

Should return `{ "status": "ok", "service": "ak-shop", ... }`.

## Routes mirrored from Express

- `POST /auth/register`, `POST /auth/login`, `GET /auth/me`
- `GET|POST /products`, `GET|PUT|DELETE /products/:id`
- `GET|POST /orders`, `GET /orders/:id`, `PATCH /orders/:id/status`
- `GET /images/:id`
- `GET /health`
