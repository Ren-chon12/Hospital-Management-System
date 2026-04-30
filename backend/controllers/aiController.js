import Groq from "groq-sdk";
import dotenv from "dotenv";

dotenv.config();

const groq = new Groq({
  apiKey: process.env.GROQ_API_KEY
});

// Analyze symptoms
export const analyzeSymptoms = async (req, res) => {
  try {
    const { symptoms } = req.body;

    if (!symptoms) {
      return res.status(400).json({
        result: "Please provide your symptoms."
      });
    }

    const completion = await groq.chat.completions.create({
      model: "llama-3.1-8b-instant",
      messages: [
        {
          role: "system",
          content:
            "You are a careful AI doctor assistant for a hospital website. Based on the symptoms, provide: 1 possible causes, 2 simple self-care tips, 3 when to consult a doctor, and 4 emergency warning signs if any. Keep the language simple. Never claim a final diagnosis. Always remind the user that this is not a substitute for professional medical advice."
        },
        {
          role: "user",
          content: `These are my symptoms: ${symptoms}`
        }
      ],
      temperature: 0.5,
      max_tokens: 500
    });

    const result =
      completion.choices?.[0]?.message?.content || "No response generated.";

    res.json({ result });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      result: "AI doctor is unavailable right now."
    });
  }
};
