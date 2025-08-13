import { createContext } from "react";
import axios from "axios";

export const ModalAIContext = createContext();

const MODAL_API_URL = "https://api.modal.com/v1/ai/query";  // Example URL, replace with actual Modal.ai endpoint
const MODAL_API_KEY = "<YOUR_MODAL_AI_API_KEY_HERE>"; // Put your Modal.ai API key here

export function ModalAIProvider({ children }) {
  // Returns formula string for natural language prompt
  async function queryAI(prompt) {
    try {
      const response = await axios.post(
        MODAL_API_URL,
        {
          prompt: prompt,
          model: "latest-excel-model" // hypothetical model name, replace with your model name or parameter
        },
        {
          headers: {
            Authorization: `Bearer ${MODAL_API_KEY}`,
            "Content-Type": "application/json"
          }
        }
      );

      // Assuming response data schema contains { formula: "some excel formula" }
      return response.data.formula || null;

    } catch (error) {
      console.error("Modal.ai query error:", error);
      throw new Error("Failed to query Modal.ai");
    }
  }

  return (
    <ModalAIContext.Provider value={{ queryAI }}>
      {children}
    </ModalAIContext.Provider>
  );
}
