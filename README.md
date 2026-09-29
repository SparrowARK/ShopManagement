# ShopManagement

Monorepo for shop tools.

## Layout

| Path | What |
|------|------|
| `Development/` | Older Python shop-data scripts |
| `mobile/` | **AK Shop** — Expo React Native app (shopkeeper + customer) |

## Mobile app (AK Shop)

```bash
cd mobile
npm install
cp .env.example .env   # then edit if needed
npx expo start
```

API is hosted on Vsys at `https://www.arkarki.com.np/api/ak-shop` (see `mobile/AK_SHOP_ON_VSYS.md`).

Local Express API (optional):

```bash
cd mobile/server
npm install
# set MONGODB_URI + JWT_SECRET in server/.env
npm run dev
```
