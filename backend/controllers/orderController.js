import Order from "../models/Order.js";
import User from "../models/User.js";
import Notification from "../models/Notification.js";

// Cart helpers
const normalizeCartItem = ({ name, price, quantity }) => ({
  name: String(name || "").trim(),
  price: Number(price),
  quantity: Number(quantity)
});

// Get cart
export const getCart = async (req, res) => {
  const user = await User.findById(req.user._id);
  res.json(user.cart || []);
};

// Add to cart
export const addToCart = async (req, res) => {
  const { name, price, quantity } = normalizeCartItem(req.body);

  if (!name || Number.isNaN(price) || Number.isNaN(quantity) || quantity <= 0) {
    return res.status(400).json({ message: "Invalid cart item data" });
  }

  const user = await User.findById(req.user._id);
  user.cart = user.cart || [];

  const existingItem = user.cart.find((item) => item.name === name);

  if (existingItem) {
    existingItem.quantity += quantity;
  } else {
    user.cart.push({ name, price, quantity });
  }

  user.markModified("cart");
  await user.save();

  res.status(201).json(user.cart);
};

// Remove one quantity from cart
export const removeFromCart = async (req, res) => {
  const { name } = req.body;
  const user = await User.findById(req.user._id);
  user.cart = user.cart || [];

  const existingItem = user.cart.find((item) => item.name === name);

  if (!existingItem) {
    return res.status(404).json({ message: "Cart item not found" });
  }

  existingItem.quantity -= 1;

  user.cart = user.cart.filter((item) => item.quantity > 0);
  user.markModified("cart");
  await user.save();

  res.json(user.cart);
};

// Place order
export const placeOrder = async (req, res) => {
  const user = await User.findById(req.user._id);

  if (!user.cart.length) {
    return res.status(400).json({ message: "Cart is empty" });
  }

  const totalAmount = user.cart.reduce(
    (sum, item) => sum + item.price * item.quantity,
    0
  );

  const order = await Order.create({
    user: user._id,
    items: user.cart,
    totalAmount,
    paymentMethod: "COD",
    paymentStatus: "pending"
  });

  user.cart = [];
  await user.save();

  await Notification.create({
    user: user._id,
    title: "Order Placed",
    message: "Your medical order has been placed with Cash on Delivery."
  });

  res.status(201).json(order);
};

// Get orders
export const getOrders = async (req, res) => {
  const { search = "" } = req.query;

  const baseQuery = req.user.role === "admin" ? {} : { user: req.user._id };
  const orders = await Order.find(baseQuery)
    .populate("user", "name email")
    .sort({ createdAt: -1 });

  const filtered = orders.filter(
    (order) =>
      order.user?.name?.toLowerCase().includes(search.toLowerCase()) ||
      order.items.some((item) =>
        item.name.toLowerCase().includes(search.toLowerCase())
      )
  );

  res.json(filtered);
};

// Update order
export const updateOrder = async (req, res) => {
  const order = await Order.findById(req.params.id);

  if (!order) {
    return res.status(404).json({ message: "Order not found" });
  }

  order.orderStatus = req.body.orderStatus || order.orderStatus;
  order.paymentStatus = req.body.paymentStatus || order.paymentStatus;
  const updated = await order.save();

  await Notification.create({
    user: updated.user,
    title: "Order Updated",
    message: `Order status is now ${updated.orderStatus}.`
  });

  res.json(updated);
};

// Delete delivered order
export const deleteOrder = async (req, res) => {
  const order = await Order.findById(req.params.id);

  if (!order) {
    return res.status(404).json({ message: "Order not found" });
  }

  if (req.user.role !== "admin" && order.user.toString() !== req.user._id.toString()) {
    return res.status(403).json({ message: "Not allowed to delete this order" });
  }

  if (order.orderStatus !== "delivered") {
    return res.status(400).json({ message: "Only delivered orders can be removed" });
  }

  await order.deleteOne();
  res.json({ message: "Order removed from history" });
};
