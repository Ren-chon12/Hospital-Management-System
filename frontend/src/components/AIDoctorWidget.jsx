import { useState } from "react";
import api from "../api/axios";

function AIDoctorWidget() {
  const [isOpen, setIsOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [symptoms, setSymptoms] = useState("");
  const [messages, setMessages] = useState([
    {
      role: "assistant",
      content:
        "Hello, I am your AI doctor assistant. Tell me your symptoms and I will suggest possible causes, self-care tips, and when to consult a doctor."
    }
  ]);

  // Send symptoms
  const handleSend = async (e) => {
    e.preventDefault();

    if (!symptoms.trim() || loading) return;

    const userMessage = symptoms.trim();
    setMessages((prev) => [...prev, { role: "user", content: userMessage }]);
    setSymptoms("");
    setLoading(true);

    try {
      const { data } = await api.post("/ai/doctor", {
        symptoms: userMessage
      });

      setMessages((prev) => [
        ...prev,
        { role: "assistant", content: data.result }
      ]);
    } catch (error) {
      setMessages((prev) => [
        ...prev,
        {
          role: "assistant",
          content:
            error.response?.data?.result || "AI doctor is unavailable right now."
        }
      ]);
    } finally {
      setLoading(false);
    }
  };

  return (
    <>
      <button className="ai-fab" onClick={() => setIsOpen((prev) => !prev)}>
        AI
      </button>

      {isOpen && (
        <div className="ai-widget">
          <div className="ai-widget-head">
            <div>
              <strong>AI Doctor</strong>
              <p>Symptom helper</p>
            </div>
            <button className="ai-close" onClick={() => setIsOpen(false)}>
              x
            </button>
          </div>

          <div className="ai-messages">
            {messages.map((message, index) => (
              <div
                key={`${message.role}-${index}`}
                className={
                  message.role === "user" ? "ai-message own" : "ai-message"
                }
              >
                <p>{message.content}</p>
              </div>
            ))}
          </div>

          <form className="ai-form" onSubmit={handleSend}>
            <textarea
              placeholder="Describe your symptoms here..."
              value={symptoms}
              onChange={(e) => setSymptoms(e.target.value)}
              rows={3}
            />
            <button className="primary-btn" type="submit" disabled={loading}>
              {loading ? "Thinking..." : "Ask AI Doctor"}
            </button>
          </form>
        </div>
      )}
    </>
  );
}

export default AIDoctorWidget;
