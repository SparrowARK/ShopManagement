import React, { useCallback, useState } from "react";
import {
  View,
  Text,
  StyleSheet,
  TextInput,
  TouchableOpacity,
  ScrollView,
  Alert,
  ActivityIndicator,
  Image,
} from "react-native";
import { SafeAreaView } from "react-native-safe-area-context";
import { useFocusEffect } from "expo-router";
import * as ImagePicker from "expo-image-picker";
import { File } from "expo-file-system";
import * as SecureStore from "expo-secure-store";
import { upload } from "../../lib/api";
import { useAuth } from "../../lib/auth";

const DEFAULT_OPENAI_MODEL = "gpt-4o-mini";
const DEFAULT_GEMINI_MODEL = "gemini-2.0-flash";

type Provider = "openai" | "gemini";

export default function AddProductScreen() {
  const [name, setName] = useState("");
  const [description, setDescription] = useState("");
  const [price, setPrice] = useState("");
  const [stock, setStock] = useState("");
  const [category, setCategory] = useState("");
  const [imageUri, setImageUri] = useState<string | null>(null);
  const [imageBase64, setImageBase64] = useState<string | null>(null);
  const [imageMime, setImageMime] = useState("image/jpeg");
  const [loading, setLoading] = useState(false);
  const [aiLoading, setAiLoading] = useState(false);
  const [provider, setProvider] = useState<Provider>("openai");
  const [openaiKey, setOpenaiKey] = useState("");
  const [openaiModel, setOpenaiModel] = useState(DEFAULT_OPENAI_MODEL);
  const [geminiKey, setGeminiKey] = useState("");
  const [geminiModel, setGeminiModel] = useState(DEFAULT_GEMINI_MODEL);
  const { user } = useAuth();

  useFocusEffect(
    useCallback(() => {
      const loadSettings = async () => {
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

      loadSettings();
    }, []),
  );

  const extractJson = (text: string) => {
    const start = text.indexOf("{");
    const end = text.lastIndexOf("}");
    if (start === -1 || end === -1) return null;
    const jsonText = text.slice(start, end + 1);
    try {
      return JSON.parse(jsonText);
    } catch {
      return null;
    }
  };

  const pickImage = async (fromCamera: boolean) => {
    const permission = fromCamera
      ? await ImagePicker.requestCameraPermissionsAsync()
      : await ImagePicker.requestMediaLibraryPermissionsAsync();

    if (!permission.granted) {
      Alert.alert("Permission needed", "Please allow access to continue");
      return;
    }

    const result = fromCamera
      ? await ImagePicker.launchCameraAsync({
          quality: 0.6,
          base64: true,
        })
      : await ImagePicker.launchImageLibraryAsync({
          quality: 0.6,
          base64: true,
        });

    if (result.canceled || !result.assets?.length) return;

    const asset = result.assets[0];
    setImageUri(asset.uri);
    setImageBase64(asset.base64 ?? null);
    setImageMime(asset.mimeType ?? "image/jpeg");
  };

  const handleAutoFill = async () => {
    if (!imageBase64) {
      Alert.alert("Add a photo", "Take or choose a product photo first");
      return;
    }

    const apiKey = provider === "openai" ? openaiKey : geminiKey;
    if (!apiKey) {
      Alert.alert("API key missing", "Add your API key in Settings");
      return;
    }

    setAiLoading(true);
    try {
      const prompt =
        "You are a retail catalog assistant. Analyze the product photo and return ONLY valid JSON with keys: name, description, price, category, stock. " +
        "Use null for unknown values. price must be a number, stock must be an integer or null. Do not include any extra text.";

      let responseText = "";

      if (provider === "openai") {
        const response = await fetch(
          "https://api.openai.com/v1/chat/completions",
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              Authorization: `Bearer ${apiKey}`,
            },
            body: JSON.stringify({
              model: openaiModel || DEFAULT_OPENAI_MODEL,
              temperature: 0.2,
              messages: [
                { role: "system", content: "Return only JSON. No commentary." },
                {
                  role: "user",
                  content: [
                    { type: "text", text: prompt },
                    {
                      type: "image_url",
                      image_url: {
                        url: `data:${imageMime};base64,${imageBase64}`,
                      },
                    },
                  ],
                },
              ],
            }),
          },
        );

        const data = await response.json();
        responseText = data?.choices?.[0]?.message?.content ?? "";
      } else {
        const response = await fetch(
          `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel || DEFAULT_GEMINI_MODEL}:generateContent?key=${apiKey}`,
          {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              contents: [
                {
                  role: "user",
                  parts: [
                    { text: prompt },
                    {
                      inlineData: {
                        mimeType: imageMime,
                        data: imageBase64,
                      },
                    },
                  ],
                },
              ],
              generationConfig: { temperature: 0.2 },
            }),
          },
        );

        const data = await response.json();
        responseText = data?.candidates?.[0]?.content?.parts?.[0]?.text ?? "";
      }

      const parsed = extractJson(responseText);
      if (!parsed) {
        throw new Error("AI response was not valid JSON");
      }

      if (parsed.name) setName(String(parsed.name));
      if (parsed.description !== undefined)
        setDescription(parsed.description ? String(parsed.description) : "");
      if (parsed.category !== undefined)
        setCategory(parsed.category ? String(parsed.category) : "");
      if (
        parsed.price !== undefined &&
        parsed.price !== null &&
        !Number.isNaN(Number(parsed.price))
      ) {
        setPrice(String(parsed.price));
      }
      if (parsed.stock !== undefined) {
        setStock(
          parsed.stock === null || parsed.stock === undefined
            ? ""
            : String(parsed.stock),
        );
      }
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : "Failed to auto-fill";
      Alert.alert("AI error", message);
    } finally {
      setAiLoading(false);
    }
  };

  const handleSubmit = async () => {
    if (!name.trim() || !price.trim()) {
      Alert.alert("Required fields", "Please fill in product name and price");
      return;
    }

    const priceNum = parseFloat(price);
    if (isNaN(priceNum) || priceNum <= 0) {
      Alert.alert("Invalid price", "Please enter a valid price");
      return;
    }

    setLoading(true);
    try {
      const formData = new FormData();
      formData.append("name", name.trim());
      formData.append("price", String(priceNum));
      if (description.trim()) formData.append("description", description.trim());
      if (category.trim()) formData.append("category", category.trim());

      const stockParsed = parseFloat(stock.trim());
      if (stock.trim() && Number.isFinite(stockParsed)) {
        formData.append("stock", String(Math.floor(Math.abs(stockParsed))));
      }

      // Attach image via expo-file-system File (required by expo/fetch FormData)
      if (imageUri) {
        const file = new File(imageUri);
        formData.append("image", file as unknown as Blob);
      }

      await upload("/products", formData);

      Alert.alert("Success", "Product added successfully", [
        {
          text: "OK",
          onPress: () => {
            setName("");
            setDescription("");
            setPrice("");
            setStock("");
            setCategory("");
            setImageUri(null);
            setImageBase64(null);
          },
        },
      ]);
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : "Failed to add product";
      Alert.alert("Error", message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <SafeAreaView style={styles.container} edges={["top"]}>
      <View style={styles.header}>
        <Text style={styles.title}>Add Product</Text>
        <Text style={styles.subtitle}>
          {user?.shop_name || "Shopkeeper view"}
        </Text>
      </View>

      <ScrollView contentContainerStyle={styles.content}>
        <View style={styles.field}>
          <Text style={styles.label}>Product Photo</Text>
          <View style={styles.photoCard}>
            {imageUri ? (
              <Image source={{ uri: imageUri }} style={styles.photoImage} />
            ) : (
              <View style={styles.photoPlaceholder}>
                <Text style={styles.photoPlaceholderText}>
                  Add a clear product photo
                </Text>
              </View>
            )}
          </View>
          <View style={styles.photoActions}>
            <TouchableOpacity
              style={styles.secondaryButton}
              onPress={() => pickImage(true)}
            >
              <Text style={styles.secondaryButtonText}>Take Photo</Text>
            </TouchableOpacity>
            <TouchableOpacity
              style={styles.secondaryButton}
              onPress={() => pickImage(false)}
            >
              <Text style={styles.secondaryButtonText}>Choose Photo</Text>
            </TouchableOpacity>
          </View>
        </View>

        <View style={styles.field}>
          <Text style={styles.label}>AI Auto-fill</Text>
          <Text style={styles.note}>
            Uses your photo to draft product details. You can edit before
            publishing.
          </Text>
          <TouchableOpacity
            style={[styles.button, aiLoading && styles.buttonDisabled]}
            onPress={handleAutoFill}
            disabled={aiLoading}
          >
            {aiLoading ? (
              <ActivityIndicator color="#FFFFFF" />
            ) : (
              <Text style={styles.buttonText}>
                Auto-fill with {provider === "openai" ? "OpenAI" : "Gemini"}
              </Text>
            )}
          </TouchableOpacity>
        </View>

        <View style={styles.field}>
          <Text style={styles.label}>Product Name *</Text>
          <TextInput
            style={styles.input}
            value={name}
            onChangeText={setName}
            placeholder="Enter product name"
            placeholderTextColor="#9E9E9E"
          />
        </View>

        <View style={styles.field}>
          <Text style={styles.label}>Description</Text>
          <TextInput
            style={[styles.input, styles.textArea]}
            value={description}
            onChangeText={setDescription}
            placeholder="Enter description (optional)"
            placeholderTextColor="#9E9E9E"
            multiline
            numberOfLines={4}
          />
        </View>

        <View style={styles.row}>
          <View style={[styles.field, styles.half]}>
            <Text style={styles.label}>Price (₹) *</Text>
            <TextInput
              style={styles.input}
              value={price}
              onChangeText={setPrice}
              placeholder="0.00"
              placeholderTextColor="#9E9E9E"
              keyboardType="decimal-pad"
            />
          </View>

          <View style={[styles.field, styles.half]}>
            <Text style={styles.label}>Stock</Text>
            <TextInput
              style={styles.input}
              value={stock}
              onChangeText={setStock}
              placeholder="Optional"
              placeholderTextColor="#9E9E9E"
              keyboardType="number-pad"
            />
          </View>
        </View>

        <View style={styles.field}>
          <Text style={styles.label}>Category</Text>
          <TextInput
            style={styles.input}
            value={category}
            onChangeText={setCategory}
            placeholder="e.g. Electronics, Clothing (optional)"
            placeholderTextColor="#9E9E9E"
          />
        </View>

        <TouchableOpacity
          style={[styles.button, loading && styles.buttonDisabled]}
          onPress={handleSubmit}
          disabled={loading}
        >
          {loading ? (
            <ActivityIndicator color="#FFFFFF" />
          ) : (
            <Text style={styles.buttonText}>Publish Product</Text>
          )}
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
  field: {
    marginBottom: 16,
  },
  label: {
    fontSize: 14,
    fontWeight: "600",
    color: "#3E2723",
    marginBottom: 8,
  },
  note: {
    fontSize: 12,
    color: "#757575",
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
  },
  photoCard: {
    borderWidth: 1,
    borderColor: "#E0E0E0",
    borderRadius: 12,
    overflow: "hidden",
    backgroundColor: "#FFFFFF",
  },
  photoImage: {
    width: "100%",
    height: 200,
  },
  photoPlaceholder: {
    height: 200,
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: "#F5F5F5",
  },
  photoPlaceholderText: {
    color: "#757575",
    fontSize: 13,
  },
  photoActions: {
    flexDirection: "row",
    gap: 12,
    marginTop: 12,
  },
  textArea: {
    height: 100,
    textAlignVertical: "top",
  },
  row: {
    flexDirection: "row",
    gap: 12,
  },
  half: {
    flex: 1,
  },
  button: {
    backgroundColor: "#3E2723",
    borderRadius: 8,
    padding: 16,
    alignItems: "center",
    marginTop: 8,
  },
  secondaryButton: {
    flex: 1,
    borderRadius: 8,
    padding: 12,
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
  buttonDisabled: {
    opacity: 0.6,
  },
  buttonText: {
    color: "#FFFFFF",
    fontSize: 16,
    fontWeight: "600",
  },
});
