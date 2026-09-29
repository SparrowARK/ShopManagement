import * as SecureStore from "expo-secure-store";
import { fetch as expoFetch } from "expo/fetch";

const API_URL = process.env.EXPO_PUBLIC_API_URL || "http://localhost:3000/api";

async function getToken(): Promise<string | null> {
  return SecureStore.getItemAsync("auth_token");
}

async function authHeaders(): Promise<Record<string, string>> {
  const token = await getToken();
  const headers: Record<string, string> = {};
  if (token) headers["Authorization"] = `Bearer ${token}`;
  return headers;
}

async function parseJson(res: Response): Promise<any> {
  const text = await res.text();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    throw new Error(
      text.slice(0, 180) || `Request failed (${res.status})`
    );
  }
}

async function request<T>(
  method: string,
  path: string,
  body?: unknown
): Promise<T> {
  const headers = await authHeaders();
  headers["Content-Type"] = "application/json";

  const res = await fetch(`${API_URL}${path}`, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });

  const data = await parseJson(res);
  if (!res.ok) {
    throw new Error(data.error || `Request failed (${res.status})`);
  }
  return data as T;
}

export function get<T>(path: string): Promise<T> {
  return request<T>("GET", path);
}

export function post<T>(path: string, body: unknown): Promise<T> {
  return request<T>("POST", path, body);
}

export function put<T>(path: string, body: unknown): Promise<T> {
  return request<T>("PUT", path, body);
}

export function patch<T>(path: string, body: unknown): Promise<T> {
  return request<T>("PATCH", path, body);
}

export function del<T>(path: string): Promise<T> {
  return request<T>("DELETE", path);
}

/**
 * Multipart upload — must use expo/fetch + expo-file-system File parts.
 * RN { uri, name, type } descriptors throw "Unsupported FormDataPart".
 */
export async function upload<T>(path: string, formData: FormData): Promise<T> {
  const headers = await authHeaders();

  const res = await expoFetch(`${API_URL}${path}`, {
    method: "POST",
    headers,
    body: formData,
  });

  const data = await parseJson(res as unknown as Response);
  if (!res.ok) {
    throw new Error(data.error || `Upload failed (${res.status})`);
  }
  return data as T;
}

export function imageUrl(imageId: string | null | undefined): string | null {
  if (!imageId) return null;
  if (imageId.startsWith("http")) return imageId;
  return `${API_URL}/images/${imageId}`;
}
