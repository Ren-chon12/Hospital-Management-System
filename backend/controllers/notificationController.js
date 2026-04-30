import Notification from "../models/Notification.js";

// Get notifications
export const getNotifications = async (req, res) => {
  const notifications = await Notification.find({ user: req.user._id }).sort({
    createdAt: -1
  });
  res.json(notifications);
};

// Mark notification read
export const markNotificationRead = async (req, res) => {
  const notification = await Notification.findById(req.params.id);

  if (!notification) {
    return res.status(404).json({ message: "Notification not found" });
  }

  notification.read = true;
  const updated = await notification.save();
  res.json(updated);
};
