import express from "express";
import {
  getContacts,
  getUsers,
  updateProfile
} from "../controllers/userController.js";
import { adminOnly, protect } from "../middleware/authMiddleware.js";

const router = express.Router();

router.get("/contacts", protect, getContacts);
router.get("/", protect, adminOnly, getUsers);
router.put("/profile", protect, updateProfile);

export default router;
