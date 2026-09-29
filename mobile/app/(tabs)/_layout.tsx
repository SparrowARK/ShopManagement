import { Tabs } from "expo-router";
import { Text, StyleSheet } from "react-native";
import { useAuth } from "../../lib/auth";

export default function TabLayout() {
  const { isShopkeeper } = useAuth();

  return (
    <Tabs
      screenOptions={{
        tabBarActiveTintColor: "#3E2723",
        tabBarInactiveTintColor: "#9E9E9E",
        headerShown: false,
        tabBarStyle: {
          borderTopWidth: 1,
          borderTopColor: "#E0E0E0",
          backgroundColor: "#FFFFFF",
        },
      }}
    >
      <Tabs.Screen
        name="index"
        options={{
          title: "Products",
          tabBarIcon: () => <Text style={styles.icon}>🛍️</Text>,
        }}
      />
      <Tabs.Screen
        name="add"
        options={{
          title: "Add Product",
          href: isShopkeeper ? undefined : null,
          tabBarIcon: () => <Text style={styles.icon}>➕</Text>,
        }}
      />
      <Tabs.Screen
        name="orders"
        options={{
          title: isShopkeeper ? "Orders" : "My Orders",
          tabBarIcon: () => <Text style={styles.icon}>📋</Text>,
        }}
      />
      <Tabs.Screen
        name="settings"
        options={{
          title: "Settings",
          tabBarIcon: () => <Text style={styles.icon}>⚙️</Text>,
        }}
      />
    </Tabs>
  );
}

const styles = StyleSheet.create({
  icon: { fontSize: 20 },
});
