import React, { useEffect, useState } from "react";
import {
  View,
  Text,
  StyleSheet,
  TextInput,
  TouchableOpacity,
  ScrollView,
  Alert,
} from "react-native";
import { SafeAreaView } from "react-native-safe-area-context";
import * as SecureStore from "expo-secure-store";
import { useAuth } from "../../lib/auth";


const DEFAULT_OPENAI_MODEL = "gpt-4o-mini";
const DEFAULT_GEMINI_MODEL = "gemini-2.0-flash";

type Provider = "openai" | "gemini";

export default function SettingsScreen() {
  const { user, logout, isShopkeeper } = useAuth();
  const [provider, setProvider] = useState<Provider>("openai");
  const [openaiKey, setOpenaiKey] = useState("");
  const [openaiModel, setOpenaiModel] = useState(DEFAULT_OPENAI_MODEL);
  const [geminiKey, setGeminiKey] = useState("");
  const [geminiModel, setGeminiModel] = useState(DEFAULT_GEMINI_MODEL);


  useEffect(() => {
    const load = async () => {
      const savedProvider = (await SecureStore.getItemAsync(
        "ai_provider",
      )) as Provider | null;
      const savedOpenaiKey = await SecureStore.getItemAsync("openai_key");
      const savedOpenaiModel = await SecureStore.getItemAsync("openai_model");
      const savedGeminiKey = await SecureStore.getItemAsync("gemini_key");
      const savedGeminiModel = await SecureStore.getItemAsync("gemini_model");

      if (savedProvider) setProvider(savedProvider);
      if (savedOpenaiKey) setOpenaiKey(savedOpenaiKey);
      if (savedOpenaiModel) setOpenaiModel(savedOpenaiModel);
      if (savedGeminiKey) setGeminiKey(savedGeminiKey);
      if (savedGeminiModel) setGeminiModel(savedGeminiModel);
    };

    load();
  }, []);

  const handleSave = async () => {
    await SecureStore.setItemAsync("ai_provider", provider);
    await SecureStore.setItemAsync("openai_key", openaiKey.trim());
    await SecureStore.setItemAsync(
      "openai_model",
      openaiModel.trim() || DEFAULT_OPENAI_MODEL,
    );
    await SecureStore.setItemAsync("gemini_key", geminiKey.trim());
    await SecureStore.setItemAsync(
      "gemini_model",
      geminiModel.trim() || DEFAULT_GEMINI_MODEL,
    );

    Alert.alert("Saved", "AI settings updated");
  };

  const handleClear = async () => {
    await SecureStore.deleteItemAsync("openai_key");
    await SecureStore.deleteItemAsync("gemini_key");
    Alert.alert("Cleared", "API keys removed from this device");
    setOpenaiKey("");
    setGeminiKey("");
  };

  const handleLogout = () => {
    Alert.alert("Log Out", "Are you sure you want to log out?", [
      { text: "Cancel", style: "cancel" },
      { text: "Log Out", style: "destructive", onPress: logout },
    ]);
  };

  return (
    <SafeAreaView style={styles.container} edges={["top"]}>
      <View style={styles.header}>
        <Text style={styles.title}>Settings</Text>
        <Text style={styles.subtitle}>{user?.name || "Settings"}</Text>
      </View>


      <ScrollView contentContainerStyle={styles.content}>
        {/* Account info */}
        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Account</Text>
          <Text style={styles.infoText}>📧 {user?.email}</Text>
          <Text style={styles.infoText}>👤 {isShopkeeper ? "Shopkeeper" : "Customer"}{user?.shop_name ? ` — ${user.shop_name}` : ""}</Text>
        </View>


        <View style={styles.section}>
          <Text style={styles.sectionTitle}>AI Provider</Text>
          <View style={styles.toggleRow}>
            <TouchableOpacity
              style={[
                styles.toggle,
                provider === "openai" && styles.toggleActive,
              ]}
              onPress={() => setProvider("openai")}
            >
              <Text
                style={[
                  styles.toggleText,
                  provider === "openai" && styles.toggleTextActive,
                ]}
              >
                OpenAI
              </Text>
            </TouchableOpacity>
            <TouchableOpacity
              style={[
                styles.toggle,
                provider === "gemini" && styles.toggleActive,
              ]}
              onPress={() => setProvider("gemini")}
            >
              <Text
                style={[
                  styles.toggleText,
                  provider === "gemini" && styles.toggleTextActive,
                ]}
              >
                Gemini
              </Text>
            </TouchableOpacity>
          </View>
        </View>

        <View style={styles.section}>
          <Text style={styles.sectionTitle}>OpenAI</Text>
          <Text style={styles.label}>API Key</Text>
          <TextInput
            style={styles.input}
            value={openaiKey}
            onChangeText={setOpenaiKey}
            placeholder="sk-..."
            placeholderTextColor="#9E9E9E"
            autoCapitalize="none"
          />
          <Text style={styles.label}>Model</Text>
          <TextInput
            style={styles.input}
            value={openaiModel}
            onChangeText={setOpenaiModel}
            placeholder={DEFAULT_OPENAI_MODEL}
            placeholderTextColor="#9E9E9E"
            autoCapitalize="none"
          />
        </View>

        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Gemini</Text>
          <Text style={styles.label}>API Key</Text>
          <TextInput
            style={styles.input}
            value={geminiKey}
            onChangeText={setGeminiKey}
            placeholder="AIza..."
            placeholderTextColor="#9E9E9E"
            autoCapitalize="none"
          />
          <Text style={styles.label}>Model</Text>
          <TextInput
            style={styles.input}
            value={geminiModel}
            onChangeText={setGeminiModel}
            placeholder={DEFAULT_GEMINI_MODEL}
            placeholderTextColor="#9E9E9E"
            autoCapitalize="none"
          />
        </View>

        <Text style={styles.note}>
          Keys are stored locally on this device. You can change or clear them
          anytime.
        </Text>

        <TouchableOpacity style={styles.primaryButton} onPress={handleSave}>
          <Text style={styles.primaryButtonText}>Save Settings</Text>
        </TouchableOpacity>

        <TouchableOpacity style={styles.secondaryButton} onPress={handleClear}>
          <Text style={styles.secondaryButtonText}>Clear API Keys</Text>
        </TouchableOpacity>

        <TouchableOpacity style={styles.logoutButton} onPress={handleLogout}>
          <Text style={styles.logoutButtonText}>Log Out</Text>
        </TouchableOpacity>
      </ScrollView>
    </SafeAreaView>

  );
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: "#FAFAFA",
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
  content: {
    padding: 16,
  },
  section: {
    backgroundColor: "#FFFFFF",
    borderRadius: 12,
    padding: 16,
    borderWidth: 1,
    borderColor: "#E0E0E0",
    marginBottom: 16,
  },
  sectionTitle: {
    fontSize: 16,
    fontWeight: "700",
    color: "#3E2723",
    marginBottom: 12,
  },
  toggleRow: {
    flexDirection: "row",
    gap: 12,
  },
  toggle: {
    flex: 1,
    paddingVertical: 12,
    borderRadius: 8,
    borderWidth: 1,
    borderColor: "#E0E0E0",
    alignItems: "center",
    backgroundColor: "#FFFFFF",
  },
  toggleActive: {
    backgroundColor: "#3E2723",
    borderColor: "#3E2723",
  },
  toggleText: {
    fontSize: 14,
    fontWeight: "600",
    color: "#3E2723",
  },
  toggleTextActive: {
    color: "#FFFFFF",
  },
  label: {
    fontSize: 13,
    fontWeight: "600",
    color: "#3E2723",
    marginBottom: 8,
  },
  input: {
    backgroundColor: "#FFFFFF",
    borderWidth: 1,
    borderColor: "#E0E0E0",
    borderRadius: 8,
    padding: 12,
    fontSize: 16,
    color: "#3E2723",
    marginBottom: 12,
  },
  note: {
    fontSize: 12,
    color: "#757575",
    marginBottom: 16,
  },
  primaryButton: {
    backgroundColor: "#3E2723",
    borderRadius: 8,
    padding: 16,
    alignItems: "center",
  },
  primaryButtonText: {
    color: "#FFFFFF",
    fontSize: 16,
    fontWeight: "600",
  },
  secondaryButton: {
    marginTop: 12,
    borderRadius: 8,
    padding: 14,
    alignItems: "center",
    borderWidth: 1,
    borderColor: "#E0E0E0",
    backgroundColor: "#FFFFFF",
  },
  secondaryButtonText: {
    color: "#3E2723",
    fontSize: 14,
    fontWeight: "600",
  },
  infoText: {
    fontSize: 14,
    color: "#757575",
    marginBottom: 6,
  },
  logoutButton: {
    marginTop: 24,
    borderRadius: 8,
    padding: 16,
    alignItems: "center",
    borderWidth: 1,
    borderColor: "#F44336",
    backgroundColor: "#FFFFFF",
  },
  logoutButtonText: {
    color: "#F44336",
    fontSize: 16,
    fontWeight: "600",
  },
});

