import { useState } from "react";
import { Link, useNavigate } from "react-router-dom";
import { useAuth } from "../context/AuthContext";

function LoginPage() {
  const navigate = useNavigate();
  const { login, sendOtp, verifyOtp, setUser, loading } = useAuth();
  const [form, setForm] = useState({ email: "", password: "" });
  const [otp, setOtp] = useState("");
  const [otpSent, setOtpSent] = useState(false);
  const [message, setMessage] = useState("");

  const handleChange = (e) => {
    setForm({ ...form, [e.target.name]: e.target.value });
  };

  // Login
  const handleLogin = async (e) => {
    e.preventDefault();
    try {
      await login(form);
      navigate("/dashboard");
    } catch (error) {
      setMessage(error.response?.data?.message || "Login failed");
    }
  };

  // Send OTP
  const handleSendOtp = async () => {
    try {
      const data = await sendOtp(form.email);
      setOtpSent(true);
      setMessage(`OTP sent. Demo OTP: ${data.demoOtp}`);
    } catch (error) {
      setMessage(error.response?.data?.message || "OTP request failed");
    }
  };

  // Verify OTP
  const handleVerifyOtp = async () => {
    try {
      const { data } = await verifyOtp(form.email, otp);
      setUser(data);
      navigate("/dashboard");
    } catch (error) {
      setMessage(error.response?.data?.message || "OTP verification failed");
    }
  };

  return (
    <div className="auth-page">
      <div className="auth-card">
        <span className="auth-badge">Smart Hospital Access</span>
        <h1>Smart Hospital Login</h1>
        <p className="auth-intro">
          Sign in to manage appointments, orders, reports, and communication.
        </p>

        <form onSubmit={handleLogin} className="stack-form">
          <input
            name="email"
            type="email"
            placeholder="Email"
            value={form.email}
            onChange={handleChange}
            required
          />
          <input
            name="password"
            type="password"
            placeholder="Password"
            value={form.password}
            onChange={handleChange}
            required
          />
          <button className="primary-btn" type="submit" disabled={loading}>
            {loading ? "Please wait..." : "Login"}
          </button>
        </form>

        <div className="otp-box">
          <button className="secondary-btn" onClick={handleSendOtp}>
            Send Email OTP
          </button>
          {otpSent && (
            <>
              <input
                type="text"
                placeholder="Enter OTP"
                value={otp}
                onChange={(e) => setOtp(e.target.value)}
              />
              <button className="secondary-btn" onClick={handleVerifyOtp}>
                Verify OTP
              </button>
            </>
          )}
        </div>

        {message && <p className="info-text">{message}</p>}
        <div className="auth-demo">
          <strong>Demo logins</strong>
          <p>Admin: admin@hospital.com / 123456</p>
          <p>Doctor: doctor@hospital.com / 123456</p>
          <p>Patient: patient@hospital.com / 123456</p>
        </div>
        <p className="auth-switch">
          New user? <Link to="/register">Create account</Link>
        </p>
      </div>
    </div>
  );
}

export default LoginPage;
