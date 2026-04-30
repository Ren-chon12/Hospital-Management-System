import express from "express";
import {
  createAppointment,
  deleteAppointment,
  getAppointments,
  updateAppointment
} from "../controllers/appointmentController.js";
import { adminOnly, protect } from "../middleware/authMiddleware.js";

const router = express.Router();

router.route("/").get(protect, getAppointments).post(protect, createAppointment);
router
  .route("/:id")
  .put(protect, updateAppointment)
  .delete(protect, adminOnly, deleteAppointment);

export default router;
