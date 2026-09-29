import mongoose, { Schema, Document } from "mongoose";

export interface IOrderItem {
  product_id: mongoose.Types.ObjectId;
  name: string;
  price: number;
  quantity: number;
}

export interface IOrder extends Document {
  _id: mongoose.Types.ObjectId;
  customer_id: mongoose.Types.ObjectId;
  shopkeeper_id: mongoose.Types.ObjectId;
  items: IOrderItem[];
  total: number;
  status: "pending" | "confirmed" | "ready" | "delivered" | "cancelled";
  customer_note?: string;
  created_at: Date;
  updated_at: Date;
}

const orderItemSchema = new Schema<IOrderItem>(
  {
    product_id: {
      type: Schema.Types.ObjectId,
      ref: "Product",
      required: true,
    },
    name: {
      type: String,
      required: true,
    },
    price: {
      type: Number,
      required: true,
      min: 0,
    },
    quantity: {
      type: Number,
      required: true,
      min: 1,
    },
  },
  { _id: false }
);

const orderSchema = new Schema<IOrder>({
  customer_id: {
    type: Schema.Types.ObjectId,
    ref: "User",
    required: true,
    index: true,
  },
  shopkeeper_id: {
    type: Schema.Types.ObjectId,
    ref: "User",
    required: true,
    index: true,
  },
  items: {
    type: [orderItemSchema],
    required: true,
    validate: {
      validator: (v: IOrderItem[]) => v.length > 0,
      message: "Order must have at least one item",
    },
  },
  total: {
    type: Number,
    required: true,
    min: 0,
  },
  status: {
    type: String,
    enum: ["pending", "confirmed", "ready", "delivered", "cancelled"],
    default: "pending",
    index: true,
  },
  customer_note: {
    type: String,
    trim: true,
    default: null,
  },
  created_at: {
    type: Date,
    default: Date.now,
  },
  updated_at: {
    type: Date,
    default: Date.now,
  },
});

orderSchema.pre("save", function (next) {
  this.updated_at = new Date();
  next();
});

orderSchema.pre("findOneAndUpdate", function (next) {
  this.set({ updated_at: new Date() });
  next();
});

orderSchema.set("toJSON", {
  transform(_doc, ret) {
    const obj = ret as unknown as Record<string, unknown>;
    delete obj.__v;
    return obj;
  },
});


export const Order = mongoose.model<IOrder>("Order", orderSchema);
