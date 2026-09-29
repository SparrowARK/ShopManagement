import React, { useCallback, useState } from "react";
import {
  View,
  Text,
  StyleSheet,
  FlatList,
  TouchableOpacity,
  ActivityIndicator,
  RefreshControl,
  Alert,
} from "react-native";
import { SafeAreaView } from "react-native-safe-area-context";
import { useFocusEffect } from "expo-router";
import { get, patch } from "../../lib/api";
import { useAuth } from "../../lib/auth";
import type { Order } from "../../lib/types";

const STATUS_COLORS: Record<string, string> = {
  pending: "#FF9800",
  confirmed: "#2196F3",
  ready: "#9C27B0",
  delivered: "#4CAF50",
  cancelled: "#F44336",
};

const STATUS_LABELS: Record<string, string> = {
  pending: "⏳ Pending",
  confirmed: "✅ Confirmed",
  ready: "📦 Ready",
  delivered: "🚚 Delivered",
  cancelled: "❌ Cancelled",
};

const NEXT_STATUS: Record<string, string> = {
  pending: "confirmed",
  confirmed: "ready",
  ready: "delivered",
};

const NEXT_ACTION_LABEL: Record<string, string> = {
  pending: "Confirm Order",
  confirmed: "Mark Ready",
  ready: "Mark Delivered",
};

export default function OrdersScreen() {
  const [orders, setOrders] = useState<Order[]>([]);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const { isShopkeeper } = useAuth();

  const fetchOrders = useCallback(async () => {
    try {
      const { orders: data } = await get<{ orders: Order[] }>("/orders");
      setOrders(data || []);
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      console.error("Error fetching orders:", message);
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  }, []);

  useFocusEffect(
    useCallback(() => {
      fetchOrders();
    }, [fetchOrders])
  );

  const onRefresh = () => {
    setRefreshing(true);
    fetchOrders();
  };

  const handleUpdateStatus = (order: Order, newStatus: string) => {
    Alert.alert(
      "Update Status",
      `Change order to "${STATUS_LABELS[newStatus]}"?`,
      [
        { text: "Cancel", style: "cancel" },
        {
          text: "Update",
          onPress: async () => {
            try {
              await patch(`/orders/${order._id}/status`, {
                status: newStatus,
              });
              fetchOrders();
            } catch (err: unknown) {
              const message =
                err instanceof Error ? err.message : "Failed to update";
              Alert.alert("Error", message);
            }
          },
        },
      ]
    );
  };

  const handleCancel = (order: Order) => {
    Alert.alert("Cancel Order", "Are you sure you want to cancel this order?", [
      { text: "No", style: "cancel" },
      {
        text: "Yes, Cancel",
        style: "destructive",
        onPress: async () => {
          try {
            await patch(`/orders/${order._id}/status`, {
              status: "cancelled",
            });
            fetchOrders();
          } catch (err: unknown) {
            const message =
              err instanceof Error ? err.message : "Failed to cancel";
            Alert.alert("Error", message);
          }
        },
      },
    ]);
  };

  const formatDate = (dateStr: string) => {
    const d = new Date(dateStr);
    return d.toLocaleDateString("en-IN", {
      day: "numeric",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  };

  const getCustomerName = (order: Order): string => {
    if (typeof order.customer_id === "object" && order.customer_id?.name) {
      return order.customer_id.name;
    }
    return "Customer";
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
        <Text style={styles.title}>
          {isShopkeeper ? "Orders" : "My Orders"}
        </Text>
        <Text style={styles.subtitle}>
          {isShopkeeper ? "Manage customer orders" : "Track your orders"}
        </Text>
      </View>

      <FlatList
        data={orders}
        keyExtractor={(item) => item._id}
        contentContainerStyle={styles.list}
        refreshControl={
          <RefreshControl refreshing={refreshing} onRefresh={onRefresh} />
        }
        ListEmptyComponent={
          <View style={styles.empty}>
            <Text style={styles.emptyText}>No orders yet</Text>
            <Text style={styles.emptySub}>
              {isShopkeeper
                ? "Orders from customers will appear here"
                : "Browse products and place an order"}
            </Text>
          </View>
        }
        renderItem={({ item }) => {
          const statusColor = STATUS_COLORS[item.status] || "#757575";
          const nextStatus = NEXT_STATUS[item.status];

          return (
            <View style={styles.card}>
              <View style={styles.cardHeader}>
                <View>
                  {isShopkeeper && (
                    <Text style={styles.customerName}>
                      {getCustomerName(item)}
                    </Text>
                  )}
                  <Text style={styles.date}>
                    {formatDate(item.created_at)}
                  </Text>
                </View>
                <View
                  style={[
                    styles.statusBadge,
                    { backgroundColor: statusColor + "20" },
                  ]}
                >
                  <Text style={[styles.statusText, { color: statusColor }]}>
                    {STATUS_LABELS[item.status] || item.status}
                  </Text>
                </View>
              </View>

              {/* Items */}
              <View style={styles.itemsList}>
                {item.items.map((orderItem, idx) => (
                  <View key={idx} style={styles.itemRow}>
                    <Text style={styles.itemName}>
                      {orderItem.quantity}× {orderItem.name}
                    </Text>
                    <Text style={styles.itemPrice}>
                      ₹{(orderItem.price * orderItem.quantity).toFixed(2)}
                    </Text>
                  </View>
                ))}
              </View>

              {/* Total */}
              <View style={styles.totalRow}>
                <Text style={styles.totalLabel}>Total</Text>
                <Text style={styles.totalValue}>₹{item.total.toFixed(2)}</Text>
              </View>

              {/* Customer note */}
              {item.customer_note && (
                <Text style={styles.note}>📝 {item.customer_note}</Text>
              )}

              {/* Actions */}
              {isShopkeeper &&
                nextStatus &&
                item.status !== "cancelled" &&
                item.status !== "delivered" && (
                  <View style={styles.actions}>
                    <TouchableOpacity
                      style={styles.actionButton}
                      onPress={() => handleUpdateStatus(item, nextStatus)}
                    >
                      <Text style={styles.actionButtonText}>
                        {NEXT_ACTION_LABEL[item.status]}
                      </Text>
                    </TouchableOpacity>
                    <TouchableOpacity
                      style={styles.cancelButton}
                      onPress={() => handleCancel(item)}
                    >
                      <Text style={styles.cancelButtonText}>Cancel</Text>
                    </TouchableOpacity>
                  </View>
                )}
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
    padding: 16,
    borderWidth: 1,
    borderColor: "#E0E0E0",
  },
  cardHeader: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "flex-start",
    marginBottom: 12,
  },
  customerName: {
    fontSize: 16,
    fontWeight: "600",
    color: "#3E2723",
  },
  date: {
    fontSize: 12,
    color: "#9E9E9E",
    marginTop: 2,
  },
  statusBadge: {
    paddingHorizontal: 10,
    paddingVertical: 4,
    borderRadius: 8,
  },
  statusText: {
    fontSize: 12,
    fontWeight: "700",
  },
  itemsList: {
    borderTopWidth: 1,
    borderTopColor: "#F0F0F0",
    paddingTop: 10,
  },
  itemRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    paddingVertical: 4,
  },
  itemName: {
    fontSize: 14,
    color: "#3E2723",
    flex: 1,
  },
  itemPrice: {
    fontSize: 14,
    color: "#757575",
    fontWeight: "600",
  },
  totalRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    borderTopWidth: 1,
    borderTopColor: "#F0F0F0",
    paddingTop: 10,
    marginTop: 8,
  },
  totalLabel: {
    fontSize: 16,
    fontWeight: "700",
    color: "#3E2723",
  },
  totalValue: {
    fontSize: 16,
    fontWeight: "700",
    color: "#4CAF50",
  },
  note: {
    fontSize: 13,
    color: "#757575",
    fontStyle: "italic",
    marginTop: 8,
  },
  actions: {
    flexDirection: "row",
    gap: 10,
    marginTop: 14,
  },
  actionButton: {
    flex: 1,
    backgroundColor: "#3E2723",
    borderRadius: 8,
    paddingVertical: 12,
    alignItems: "center",
  },
  actionButtonText: {
    color: "#FFFFFF",
    fontSize: 14,
    fontWeight: "600",
  },
  cancelButton: {
    borderRadius: 8,
    paddingVertical: 12,
    paddingHorizontal: 20,
    alignItems: "center",
    borderWidth: 1,
    borderColor: "#F44336",
  },
  cancelButtonText: {
    color: "#F44336",
    fontSize: 14,
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
    textAlign: "center",
  },
});
