import User from "../models/User.js";
import Patient from "../models/Patient.js";
import Appointment from "../models/Appointment.js";
import Order from "../models/Order.js";

// Dashboard summary
export const getDashboard = async (req, res) => {
  const [users, patients, appointments, orders] = await Promise.all([
    User.countDocuments(),
    Patient.countDocuments(),
    Appointment.countDocuments(),
    Order.countDocuments()
  ]);

  res.json({ users, patients, appointments, orders });
};
