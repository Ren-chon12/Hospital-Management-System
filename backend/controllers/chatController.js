import Message from "../models/Message.js";

// Get conversations
export const getMessages = async (req, res) => {
  const { receiverId } = req.params;

  const messages = await Message.find({
    $or: [
      { sender: req.user._id, receiver: receiverId },
      { sender: receiverId, receiver: req.user._id }
    ]
  })
    .populate("sender", "name")
    .populate("receiver", "name")
    .sort({ createdAt: 1 });

  res.json(messages);
};

// Save chat message
export const saveMessage = async (req, res) => {
  const message = await Message.create({
    sender: req.user._id,
    receiver: req.body.receiver,
    text: req.body.text
  });

  const populated = await Message.findById(message._id)
    .populate("sender", "name")
    .populate("receiver", "name");

  res.status(201).json(populated);
};
