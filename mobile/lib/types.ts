export interface Product {
  _id: string;
  shopkeeper_id: string;
  name: string;
  description?: string | null;
  price: number;
  image_id?: string | null;
  category?: string | null;
  stock?: number | null;
  created_at: string;
  updated_at: string;
}

export interface User {
  _id: string;
  email: string;
  name: string;
  role: "shopkeeper" | "customer";
  shop_name?: string | null;
  created_at: string;
}

export interface OrderItem {
  product_id: string;
  name: string;
  price: number;
  quantity: number;
}

export interface Order {
  _id: string;
  customer_id: string | { _id: string; name: string; email: string };
  shopkeeper_id: string;
  items: OrderItem[];
  total: number;
  status: "pending" | "confirmed" | "ready" | "delivered" | "cancelled";
  customer_note?: string | null;
  created_at: string;
  updated_at: string;
}
