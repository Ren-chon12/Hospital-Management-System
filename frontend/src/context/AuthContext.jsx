import { createContext, useContext, useEffect, useState } from "react";
import api from "../api/axios";

const AuthContext = createContext(null);

export const AuthProvider = ({ children }) => {
  const [user, setUser] = useState(() => {
    const stored = localStorage.getItem("smartHospitalUser");
    return stored ? JSON.parse(stored) : null;
  });
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (user) {
      localStorage.setItem("smartHospitalUser", JSON.stringify(user));
    } else {
      localStorage.removeItem("smartHospitalUser");
    }
  }, [user]);

  // Auth helpers
  const login = async (payload) => {
    setLoading(true);
    try {
      const { data } = await api.post("/auth/login", payload);
      setUser(data);
      return data;
    } finally {
      setLoading(false);
    }
  };

  const register = async (payload) => {
    setLoading(true);
    try {
      const { data } = await api.post("/auth/register", payload);
      setUser(data);
      return data;
    } finally {
      setLoading(false);
    }
  };

  const sendOtp = async (email) => {
    const { data } = await api.post("/auth/send-otp", { email });
    return data;
  };

  const verifyOtp = async (email, otp) => {
    return api.post("/auth/verify-otp", { email, otp });
  };

  const logout = () => setUser(null);

  return (
    <AuthContext.Provider
      value={{ user, setUser, loading, login, register, sendOtp, verifyOtp, logout }}
    >
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => useContext(AuthContext);
