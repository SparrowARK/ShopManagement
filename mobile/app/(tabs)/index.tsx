import React, { useCallback, useState } from "react";
import {
  View,
  Text,
  StyleSheet,
  FlatList,
  Image,
  ActivityIndicator,
  RefreshControl,
  TouchableOpacity,
  Alert,
} from "react-native";
import { SafeAreaView } from "react-native-safe-area-context";
import { useFocusEffect } from "expo-router";
import { get, post, imageUrl } from "../../lib/api";
import { useAuth } from "../../lib/auth";
import type { Product } from "../../lib/types";

export default function ProductsScreen() {
  const [products, setProducts] = useState<Product[]>([]);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const { isCustomer, isShopkeeper, user } = useAuth();

  const fetchProducts = useCallback(async () => {
    try {
      const qs =
        isShopkeeper && user?._id
          ? `?shopkeeper_id=${encodeURIComponent(user._id)}`
          : "";
      const { products: data } = await get<{ products: Product[] }>(
        `/products${qs}`
      );
      setProducts(data || []);
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      console.error("Error fetching products:", message);
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  }, [isShopkeeper, user?._id]);

  useFocusEffect(
    useCallback(() => {
      fetchProducts();
    }, [fetchProducts])
  );

  const onRefresh = () => {
    setRefreshing(true);
    fetchProducts();
  };

  const handleOrder = async (product: Product) => {
    if (product.stock != null && product.stock < 1) {
      Alert.alert("Out of stock", "This product is currently unavailable");
      return;
    }

    Alert.alert(
      "Place Order",
      `Order 1× ${product.name} for ₹${Number(product.price).toFixed(2)}?`,
      [
        { text: "Cancel", style: "cancel" },
        {
          text: "Order Now",
          onPress: async () => {
            try {
              await post("/orders", {
                items: [{ product_id: product._id, quantity: 1 }],
              });
              Alert.alert("Order Placed", "The shopkeeper will be notified");
              fetchProducts();
            } catch (err: unknown) {
              const message =
                err instanceof Error ? err.message : "Failed to place order";
              Alert.alert("Order failed", message);
            }
          },
        },
      ]
    );
  };

  if (loading) {
    return (
      <SafeAreaView style={styles.container}>
        <View style={styles.center}>
          <ActivityIndicator size="large" color="#3E2723" />
        </View>
      </SafeAreaView>
    );
  }

  return (
    <SafeAreaView style={styles.container} edges={["top"]}>
      <View style={styles.header}>
        <Text style={styles.title}>AK Shop</Text>
        <Text style={styles.subtitle}>
          {isCustomer ? "Browse & Order" : "Your Products"}
        </Text>
      </View>

      <FlatList
        data={products}
        keyExtractor={(item) => item._id}
        contentContainerStyle={styles.list}
        refreshControl={
          <RefreshControl refreshing={refreshing} onRefresh={onRefresh} />
        }
        ListEmptyComponent={
          <View style={styles.empty}>
            <Text style={styles.emptyText}>No products yet</Text>
            <Text style={styles.emptySub}>
              {isCustomer
                ? "Check back soon!"
                : "Add products to get started"}
            </Text>
          </View>
        }
        renderItem={({ item }) => {
          const imgSrc = imageUrl(item.image_id);
          const outOfStock = item.stock != null && item.stock < 1;
          return (
            <View style={styles.card}>
              {imgSrc && (
                <Image source={{ uri: imgSrc }} style={styles.image} />
              )}
              <View style={styles.cardContent}>
                <Text style={styles.productName}>{item.name}</Text>
                {item.description && (
                  <Text style={styles.description}>{item.description}</Text>
                )}
                <View style={styles.footer}>
                  <Text style={styles.price}>
                    ₹
                    {item.price != null
                      ? Number(item.price).toFixed(2)
                      : "0.00"}
                  </Text>
                  {item.stock != null && (
                    <Text
                      style={[styles.stock, outOfStock && styles.stockOut]}
                    >
                      {outOfStock ? "Out of stock" : `Stock: ${item.stock}`}
                    </Text>
                  )}
                </View>
                {isCustomer && (
                  <TouchableOpacity
                    style={[
                      styles.orderButton,
                      outOfStock && styles.orderButtonDisabled,
                    ]}
                    onPress={() => handleOrder(item)}
                    disabled={outOfStock}
                  >
                    <Text style={styles.orderButtonText}>
                      {outOfStock ? "Out of Stock" : "Order"}
                    </Text>
                  </TouchableOpacity>
                )}
              </View>
            </View>
          );
        }}
      />
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: "#FAFAFA",
  },
  center: {
    flex: 1,
    justifyContent: "center",
    alignItems: "center",
  },
  header: {
    backgroundColor: "#FFFFFF",
    padding: 16,
    borderBottomWidth: 1,
    borderBottomColor: "#E0E0E0",
  },
  title: {
    fontSize: 24,
    fontWeight: "700",
    color: "#3E2723",
  },
  subtitle: {
    fontSize: 14,
    color: "#757575",
    marginTop: 4,
  },
  list: {
    padding: 16,
  },
  card: {
    backgroundColor: "#FFFFFF",
    borderRadius: 12,
    marginBottom: 12,
    overflow: "hidden",
    borderWidth: 1,
    borderColor: "#E0E0E0",
  },
  image: {
    width: "100%",
    height: 200,
    backgroundColor: "#F5F5F5",
  },
  cardContent: {
    padding: 16,
  },
  productName: {
    fontSize: 18,
    fontWeight: "600",
    color: "#3E2723",
    marginBottom: 4,
  },
  description: {
    fontSize: 14,
    color: "#757575",
    marginBottom: 12,
  },
  footer: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
  },
  price: {
    fontSize: 20,
    fontWeight: "700",
    color: "#4CAF50",
  },
  stock: {
    fontSize: 14,
    color: "#757575",
  },
  stockOut: {
    color: "#F44336",
    fontWeight: "600",
  },
  orderButton: {
    marginTop: 12,
    backgroundColor: "#3E2723",
    borderRadius: 8,
    paddingVertical: 12,
    alignItems: "center",
  },
  orderButtonDisabled: {
    backgroundColor: "#BDBDBD",
  },
  orderButtonText: {
    color: "#FFFFFF",
    fontSize: 15,
    fontWeight: "600",
  },
  empty: {
    padding: 48,
    alignItems: "center",
  },
  emptyText: {
    fontSize: 18,
    color: "#757575",
    marginBottom: 8,
  },
  emptySub: {
    fontSize: 14,
    color: "#9E9E9E",
  },
});
