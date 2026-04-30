import Patient from "../models/Patient.js";
import Notification from "../models/Notification.js";
import User from "../models/User.js";

// Create patient
export const createPatient = async (req, res) => {
  const { user: userId, patientName, age, gender, bloodGroup, disease } = req.body;
  let patientUserId = userId;

  if (!patientUserId && patientName) {
    const cleanName = patientName.trim();
    const generatedEmail = `${cleanName.toLowerCase().replace(/\s+/g, ".")}.${Date.now()}@patient.local`;

    const createdUser = await User.create({
      name: cleanName,
      email: generatedEmail,
      password: "123456",
      role: "patient"
    });

    patientUserId = createdUser._id;
  }

  if (!patientUserId) {
    return res.status(400).json({ message: "Patient name is required" });
  }

  const patient = await Patient.create({
    user: patientUserId,
    age,
    gender,
    bloodGroup,
    disease
  });

  await Notification.create({
    user: patientUserId,
    title: "Profile Created",
    message: "Your patient profile was added by the admin."
  });

  const populatedPatient = await Patient.findById(patient._id).populate(
    "user",
    "name email phone"
  );

  res.status(201).json(populatedPatient);
};

// Get patients
export const getPatients = async (req, res) => {
  const { search = "" } = req.query;
  const query = req.user.role === "patient" ? { user: req.user._id } : {};

  const patients = await Patient.find(query)
    .populate("user", "name email phone")
    .sort({ createdAt: -1 });

  const filtered = patients.filter(
    (item) =>
      item.user?.name?.toLowerCase().includes(search.toLowerCase()) ||
      item.disease?.toLowerCase().includes(search.toLowerCase())
  );

  res.json(filtered);
};

// Update patient
export const updatePatient = async (req, res) => {
  const patient = await Patient.findById(req.params.id);

  if (!patient) {
    return res.status(404).json({ message: "Patient not found" });
  }

  if (req.user.role === "patient" && patient.user.toString() !== req.user._id.toString()) {
    return res.status(403).json({ message: "Not allowed to update this patient" });
  }

  Object.assign(patient, req.body);
  const updated = await patient.save();
  res.json(updated);
};

// Delete patient
export const deletePatient = async (req, res) => {
  const patient = await Patient.findById(req.params.id);

  if (!patient) {
    return res.status(404).json({ message: "Patient not found" });
  }

  await patient.deleteOne();
  res.json({ message: "Patient deleted" });
};
