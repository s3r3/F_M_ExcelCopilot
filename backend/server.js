import express from "express";
import cors from "cors";
import fetch from "node-fetch";
import dotenv from "dotenv";
dotenv.config();

const app = express();
const PORT = 3001;
const COHERE_API_KEY = process.env.COHERE_API_KEY;

app.use(cors({ origin: "*", methods: ["POST", "GET"], allowedHeaders: ["Content-Type"] }));
app.use(express.json());

// Cek koneksi
app.get("/", (req, res) => {
  res.send("✅ Excel AI Copilot backend is running");
});

// Endpoint untuk menerima prompt dari Excel Add-in
app.post("/api/ai", async (req, res) => {
  try {
    const { command, sheetData = "", memory = [] } = req.body;
    if (!command) return res.status(400).json({ error: "Command is required" });

    // Format memori obrolan jika ada
    const chatHistory = memory.map(entry => {
      const role = entry.role === "user" ? "User" : "Assistant";
      return `${role}: ${entry.message}`;
    }).join("\n");

    const prompt = `
You are an AI assistant embedded inside Microsoft Excel. Respond in plain text what should be done in Excel based on user instructions. You are allowed to return data that should be filled in the cells.

User's instruction:
"${command}"

Excel sheet context:
${sheetData ? sheetData : "(no data provided)"}

${chatHistory ? `\nChat History:\n${chatHistory}` : ""}
Assistant:
    `.trim();

    const response = await fetch("https://api.cohere.ai/v1/chat", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${COHERE_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message: prompt,
        model: "command-r-plus",
        temperature: 0.3,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("Cohere API Error:", errorText);
      return res.status(500).json({ error: "Failed to fetch response from Cohere API" });
    }

    const result = await response.json();
    const reply = result.text || result.reply || "No response from AI.";

    res.json({ reply });
  } catch (err) {
    console.error("Server error:", err);
    res.status(500).json({ error: "Internal Server Error" });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Excel AI backend running at http://localhost:${PORT}`);
});
