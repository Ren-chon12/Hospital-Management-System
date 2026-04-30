import express from "express";
import {
  createPatient,
  deletePatient,
  getPatients,
  updatePatient
} from "../controllers/patientController.js";
import { adminOnly, protect } from "../middleware/authMiddleware.js";

const router = express.Router();

router
  .route("/")
  .get(protect, getPatients)
  .post(protect, adminOnly, createPatient);

router
  .route("/:id")
  .put(protect, adminOnly, updatePatient)
  .delete(protect, adminOnly, deletePatient);

export default router;
