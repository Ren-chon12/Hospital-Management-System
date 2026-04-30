import express from "express";
import {
  addToCart,
  deleteOrder,
  getCart,
  getOrders,
  placeOrder,
  removeFromCart,
  updateOrder
} from "../controllers/orderController.js";
import { adminOnly, protect } from "../middleware/authMiddleware.js";

const router = express.Router();

router.get("/cart", protect, getCart);
router.post("/cart", protect, addToCart);
router.put("/cart", protect, removeFromCart);
router.post("/place", protect, placeOrder);
router.get("/", protect, getOrders);
router.delete("/:id", protect, deleteOrder);
router.put("/:id", protect, adminOnly, updateOrder);

export default router;
