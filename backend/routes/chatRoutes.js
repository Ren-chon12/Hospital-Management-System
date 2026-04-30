import express from "express";
import { getMessages, saveMessage } from "../controllers/chatController.js";
import { protect } from "../middleware/authMiddleware.js";

const router = express.Router();

router.get("/:receiverId", protect, getMessages);
router.post("/", protect, saveMessage);

export default router;
