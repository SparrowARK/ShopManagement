import mongoose, { Schema, Document } from "mongoose";

export interface IProduct extends Document {
  _id: mongoose.Types.ObjectId;
  shopkeeper_id: mongoose.Types.ObjectId;
  name: string;
  description?: string;
  price: number;
  stock?: number;
  category?: string;
  image_id?: string;
  created_at: Date;
  updated_at: Date;
}

const productSchema = new Schema<IProduct>({
  shopkeeper_id: {
    type: Schema.Types.ObjectId,
    ref: "User",
    required: true,
    index: true,
  },
  name: {
    type: String,
    required: true,
    trim: true,
  },
  description: {
    type: String,
    trim: true,
    default: null,
  },
  price: {
    type: Number,
    required: true,
    min: 0,
  },
  stock: {
    type: Number,
    min: 0,
    default: null,
  },
  category: {
    type: String,
    trim: true,
    default: null,
    index: true,
  },
  image_id: {
    type: String,
    default: null,
  },
  created_at: {
    type: Date,
    default: Date.now,
    index: true,
  },
  updated_at: {
    type: Date,
    default: Date.now,
  },
});

// Auto-update updated_at on save
productSchema.pre("save", function (next) {
  this.updated_at = new Date();
  next();
});

productSchema.pre("findOneAndUpdate", function (next) {
  this.set({ updated_at: new Date() });
  next();
});

productSchema.set("toJSON", {
  transform(_doc, ret) {
    const obj = ret as unknown as Record<string, unknown>;
    delete obj.__v;
    return obj;
  },
});


export const Product = mongoose.model<IProduct>("Product", productSchema);
