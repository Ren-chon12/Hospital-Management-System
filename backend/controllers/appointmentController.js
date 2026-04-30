import Appointment from "../models/Appointment.js";
import Notification from "../models/Notification.js";

// Create appointment
export const createAppointment = async (req, res) => {
  const appointment = await Appointment.create(req.body);

  await Notification.create({
    user: req.body.patient,
    title: "Appointment Booked",
    message: `Your appointment is scheduled on ${req.body.appointmentDate} at ${req.body.appointmentTime}.`
  });

  res.status(201).json(appointment);
};

// Get appointments
export const getAppointments = async (req, res) => {
  const { status = "", search = "" } = req.query;
  const query = {};

  if (status) {
    query.status = status;
  }

  if (req.user.role === "patient") {
    query.patient = req.user._id;
  }

  if (req.user.role === "doctor") {
    query.doctor = req.user._id;
  }

  const appointments = await Appointment.find(query)
    .populate("patient", "name email")
    .populate("doctor", "name email")
    .sort({ createdAt: -1 });

  const filtered = appointments.filter(
    (item) =>
      item.patient?.name?.toLowerCase().includes(search.toLowerCase()) ||
      item.doctor?.name?.toLowerCase().includes(search.toLowerCase()) ||
      item.reason?.toLowerCase().includes(search.toLowerCase())
  );

  res.json(filtered);
};

// Update appointment
export const updateAppointment = async (req, res) => {
  const appointment = await Appointment.findById(req.params.id);

  if (!appointment) {
    return res.status(404).json({ message: "Appointment not found" });
  }

  const isOwnerPatient = appointment.patient.toString() === req.user._id.toString();
  const isOwnerDoctor = appointment.doctor.toString() === req.user._id.toString();
  const isAdmin = req.user.role === "admin";

  if (!isAdmin && !isOwnerPatient && !isOwnerDoctor) {
    return res.status(403).json({ message: "Not allowed to update this appointment" });
  }

  Object.assign(appointment, req.body);
  const updated = await appointment.save();

  await Notification.create({
    user: updated.patient,
    title: "Appointment Updated",
    message: `Appointment status changed to ${updated.status}.`
  });

  res.json(updated);
};

// Delete appointment
export const deleteAppointment = async (req, res) => {
  const appointment = await Appointment.findById(req.params.id);

  if (!appointment) {
    return res.status(404).json({ message: "Appointment not found" });
  }

  await appointment.deleteOne();
  res.json({ message: "Appointment deleted" });
};
