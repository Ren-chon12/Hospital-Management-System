import { useEffect, useMemo, useRef, useState } from "react";
import { io } from "socket.io-client";
import api from "../api/axios";

const socket = io("http://localhost:5000", { autoConnect: false });

function ChatBox({ user }) {
  const [contacts, setContacts] = useState([]);
  const [selectedUser, setSelectedUser] = useState(null);
  const [messages, setMessages] = useState([]);
  const [text, setText] = useState("");
  const bottomRef = useRef(null);

  const otherUsers = useMemo(
    () => contacts.filter((item) => item._id !== user._id),
    [contacts, user]
  );

  useEffect(() => {
    const loadContacts = async () => {
      const { data } = await api.get("/users/contacts");
      setContacts(data);
      if (data.length) {
        setSelectedUser(data[0]);
      }
    };

    loadContacts();
  }, []);

  useEffect(() => {
    socket.connect();
    socket.emit("join", user._id);
    socket.on("receive_message", (message) => {
      if (message.sender === selectedUser?._id || message.receiver === selectedUser?._id) {
        setMessages((prev) => [...prev, message]);
      }
    });

    return () => {
      socket.off("receive_message");
      socket.disconnect();
    };
  }, [selectedUser, user]);

  useEffect(() => {
    const loadMessages = async () => {
      if (!selectedUser) return;
      const { data } = await api.get(`/chat/${selectedUser._id}`);
      setMessages(data);
    };

    loadMessages();
  }, [selectedUser]);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  // Send message
  const handleSend = async (e) => {
    e.preventDefault();

    if (!text.trim() || !selectedUser) return;

    const payload = {
      receiver: selectedUser._id,
      text
    };

    const { data } = await api.post("/chat", payload);
    setMessages((prev) => [...prev, data]);
    socket.emit("send_message", {
      ...data,
      sender: user._id,
      receiver: selectedUser._id
    });
    setText("");
  };

  return (
    <div className="chat-layout">
      <div className="chat-users">
        {otherUsers.map((item) => (
          <button
            key={item._id}
            className={selectedUser?._id === item._id ? "chat-user active" : "chat-user"}
            onClick={() => setSelectedUser(item)}
          >
            {item.name} ({item.role})
          </button>
        ))}
      </div>

      <div className="chat-panel">
        <div className="messages">
          {messages.map((item) => (
            <div
              key={item._id || `${item.sender}-${item.createdAt}`}
              className={
                item.sender?._id === user._id || item.sender === user._id
                  ? "message own"
                  : "message"
              }
            >
              <p>{item.text}</p>
            </div>
          ))}
          <div ref={bottomRef} />
        </div>

        <form className="inline-form" onSubmit={handleSend}>
          <input
            type="text"
            placeholder="Type a message"
            value={text}
            onChange={(e) => setText(e.target.value)}
          />
          <button className="primary-btn" type="submit">
            Send
          </button>
        </form>
      </div>
    </div>
  );
}

export default ChatBox;
