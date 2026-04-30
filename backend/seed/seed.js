import dotenv from "dotenv";
import connectDB from "../config/db.js";
import User from "../models/User.js";
import Patient from "../models/Patient.js";
import Appointment from "../models/Appointment.js";
import Order from "../models/Order.js";
import Notification from "../models/Notification.js";
import Message from "../models/Message.js";
import { users } from "../utils/seedData.js";


dotenv.config();


const seedDatabase = async () => {
  try {
    await connectDB();

    await Promise.all([
      User.deleteMany(),
      Patient.deleteMany(),
      Appointment.deleteMany(),
      Order.deleteMany(),
      Notification.deleteMany(),
      Message.deleteMany()
    ]);

    const createdUsers = [];
    for (const user of users) {
      const createdUser = await User.create(user);
      createdUsers.push(createdUser);
    }
    const admin = createdUsers[0];
    const doctor = createdUsers[1];
    const patientUser = createdUsers[2];

    await Patient.create({
      user: patientUser._id,
      age: 28,
      gender: "Male",
      bloodGroup: "B+",
      disease: "Fever",
      reports: ["Initial blood report"],
      prescriptions: ["Paracetamol 650mg"]
    });

    await Appointment.create({
      patient: patientUser._id,
      doctor: doctor._id,
      appointmentDate: "2026-04-10",
      appointmentTime: "11:00 AM",
      reason: "Routine checkup",
      ePrescription: "Hydration and rest"
    });

    await Order.create({
      user: patientUser._id,
      items: [
        { name: "Vitamin Tablets", price: 250, quantity: 2 },
        { name: "Thermometer", price: 180, quantity: 1 }
      ],
      totalAmount: 680,
      paymentMethod: "COD",
      paymentStatus: "pending",
      orderStatus: "placed"
    });

    await Notification.insertMany([
      {
        user: admin._id,
        title: "Seed Complete",
        message: "Dummy data has been added successfully."
      },
      {
        user: patientUser._id,
        title: "Welcome",
        message: "Your patient account is ready for testing."
      }
    ]);

    await Message.create({
      sender: patientUser._id,
      receiver: doctor._id,
      text: "Hello doctor, I have uploaded my report."
    });

    console.log("Database seeded successfully");
    process.exit();
  } catch (error) {
    console.error(error);
    process.exit(1);
  }
};

seedDatabase();
