import User from "../models/User.js";

// List users
export const getUsers = async (req, res) => {
  const { role = "", search = "" } = req.query;

  const query = {
    name: { $regex: search, $options: "i" }
  };

  if (role) {
    query.role = role;
  }

  const users = await User.find(query).select("-password").sort({ createdAt: -1 });
  res.json(users);
};

// Get chat contacts
export const getContacts = async (req, res) => {
  const contacts = await User.find({ _id: { $ne: req.user._id } })
    .select("name email role")
    .sort({ name: 1 });

  res.json(contacts);
};

// Update user profile
export const updateProfile = async (req, res) => {
  const user = await User.findById(req.user._id);

  if (!user) {
    return res.status(404).json({ message: "User not found" });
  }

  user.name = req.body.name || user.name;
  user.phone = req.body.phone || user.phone;
  user.address = req.body.address || user.address;

  if (req.body.password) {
    user.password = req.body.password;
  }

  const updated = await user.save();
  res.json({
    _id: updated._id,
    name: updated.name,
    email: updated.email,
    role: updated.role,
    phone: updated.phone,
    address: updated.address
  });
};
