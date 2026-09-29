import React, { createContext, useContext, useEffect, useState } from "react";
import * as SecureStore from "expo-secure-store";
import { post, get } from "./api";

export interface User {
  _id: string;
  email: string;
  name: string;
  role: "shopkeeper" | "customer";
  shop_name?: string | null;
  created_at: string;
}

interface AuthState {
  user: User | null;
  token: string | null;
  loading: boolean;
  login: (email: string, password: string) => Promise<void>;
  register: (data: RegisterData) => Promise<void>;
  logout: () => Promise<void>;
  isShopkeeper: boolean;
  isCustomer: boolean;
  isLoggedIn: boolean;
}

interface RegisterData {
  email: string;
  password: string;
  name: string;
  role: "shopkeeper" | "customer";
  shop_name?: string;
}

interface AuthResponse {
  token: string;
  user: User;
}

const AuthContext = createContext<AuthState | undefined>(undefined);

export function AuthProvider({ children }: { children: React.ReactNode }) {
  const [user, setUser] = useState<User | null>(null);
  const [token, setToken] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  // On mount, check for existing token
  useEffect(() => {
    const restore = async () => {
      try {
        const savedToken = await SecureStore.getItemAsync("auth_token");
        if (savedToken) {
          setToken(savedToken);
          // Validate token by fetching current user
          const { user: me } = await get<{ user: User }>("/auth/me");
          setUser(me);
        }
      } catch {
        // Token invalid or expired — clear it
        await SecureStore.deleteItemAsync("auth_token");
        setToken(null);
        setUser(null);
      } finally {
        setLoading(false);
      }
    };
    restore();
  }, []);

  const login = async (email: string, password: string) => {
    const { token: newToken, user: newUser } = await post<AuthResponse>(
      "/auth/login",
      { email, password }
    );
    await SecureStore.setItemAsync("auth_token", newToken);
    setToken(newToken);
    setUser(newUser);
  };

  const register = async (data: RegisterData) => {
    const { token: newToken, user: newUser } = await post<AuthResponse>(
      "/auth/register",
      data
    );
    await SecureStore.setItemAsync("auth_token", newToken);
    setToken(newToken);
    setUser(newUser);
  };

  const logout = async () => {
    await SecureStore.deleteItemAsync("auth_token");
    setToken(null);
    setUser(null);
  };

  const value: AuthState = {
    user,
    token,
    loading,
    login,
    register,
    logout,
    isShopkeeper: user?.role === "shopkeeper",
    isCustomer: user?.role === "customer",
    isLoggedIn: !!user,
  };

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export function useAuth(): AuthState {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error("useAuth must be used within an AuthProvider");
  }
  return context;
}
