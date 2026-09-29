The migration is complete. Here's a summary:

✅ Done — Supabase fully removed, MongoDB Atlas fully set up.

What was built:
Backend server (server/) — Express.js + Mongoose + JWT + GridFS:

Auth — register/login with bcrypt-hashed passwords, JWT tokens (7-day expiry)
Products — full CRUD with image upload (photos stored in MongoDB GridFS)
Orders — customers place orders, shopkeepers manage status, stock auto-decrements
Images — served from GridFS with 24h cache headers
Mobile app — fully rewired:

Login/Register screens with customer/shopkeeper role picker
Auth-gated routing — login required, auto-redirect based on auth state
Products screen — customers see "Order" buttons, shopkeepers see "Your Products"
Orders tab — customers track orders, shopkeepers manage them (confirm → ready → delivered)
Settings — shows account info + logout button
To run:
cd server && npm run dev (starts backend on port 3000)
npx expo start (starts the mobile app)
A test shopkeeper account (aaditya@akshop.com / test123) is already created in your Atlas database