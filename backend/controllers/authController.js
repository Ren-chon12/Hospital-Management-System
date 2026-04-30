import User from "../models/User.js";
import Notification from "../models/Notification.js";
import generateToken from "../utils/generateToken.js";
import sendOtpEmail from "../utils/sendOtpEmail.js";

const buildAuthResponse = (user) => ({
  _id: user._id,
  name: user.name,
  email: user.email,
  role: user.role,
  phone: user.phone,
  address: user.address,
  token: generateToken(user._id)
});

// Register user
export const registerUser = async (req, res) => {
  const { name, email, password, role, phone, address } = req.body;

  const exists = await User.findOne({ email });
  if (exists) {
    return res.status(400).json({ message: "User already exists" });
  }

  const user = await User.create({
    name,
    email,
    password,
    role: role || "patient",
    phone,
    address
  });

  await Notification.create({
    user: user._id,
    title: "Welcome",
    message: "Your Smart Hospital account has been created."
  });

  res.status(201).json(buildAuthResponse(user));
};

// Login user
export const loginUser = async (req, res) => {
  const { email, password } = req.body;
  const user = await User.findOne({ email });

  if (!user || !(await user.matchPassword(password))) {
    return res.status(401).json({ message: "Invalid credentials" });
  }

  res.json(buildAuthResponse(user));
};

// Send email OTP
export const sendOtp = async (req, res) => {
  const { email } = req.body;
  const user = await User.findOne({ email });

  if (!user) {
    return res.status(404).json({ message: "User not found" });
  }

  const otp = `${Math.floor(100000 + Math.random() * 900000)}`;
  user.otp = otp;
  user.otpExpiresAt = new Date(Date.now() + 10 * 60 * 1000);
  user.isOtpVerified = false;
  await user.save();

  await sendOtpEmail(email, otp);

  res.json({
    message: "OTP sent successfully",
    demoOtp: otp
  });
};

// Verify email OTP
export const verifyOtp = async (req, res) => {
  const { email, otp } = req.body;
  const user = await User.findOne({ email });

  if (!user || user.otp !== otp || !user.otpExpiresAt || user.otpExpiresAt < new Date()) {
    return res.status(400).json({ message: "Invalid or expired OTP" });
  }

  user.isOtpVerified = true;
  user.otp = "";
  user.otpExpiresAt = null;
  await user.save();

  res.json(buildAuthResponse(user));
};

// Get profile
export const getProfile = async (req, res) => {
  res.json(req.user);
};
