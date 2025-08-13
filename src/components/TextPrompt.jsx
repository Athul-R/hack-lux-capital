import React, { useState, useContext } from "react";
import { ModalAIContext } from "../services/modalAIService";

export default function TextPrompt({ sheetData, setSheetData }) {
  const [prompt, setPrompt] = useState("");
  const [status, setStatus] = useState("");
  const { queryAI } = useContext(ModalAIContext);

  const handleSubmit = async () => {
    if (!prompt.trim()) return;
    setStatus("Processing...");
    try {
      // Query Modal.ai for formula generation
      const formula = await queryAI(prompt);
      if (!formula) {
        setStatus("No formula returned");
        return;
      }

      // For demonstration, insert formula in first empty cell in first row after headers
      const newData = [...sheetData];
      if (newData.length < 2) {
        newData.push([]);
      }
      if (newData[1].length === 0) {
        newData[1][0] = formula;
      } else {
        newData[1].push(formula);
      }
      setSheetData(newData);
      setStatus("Formula inserted: " + formula);
      setPrompt("");
    } catch (error) {
      setStatus("Error: " + error.message);
    }
  };

  return (
    <div className="textprompt">
      <h3>Natural Language to Excel Formula</h3>
      <textarea 
        rows="3" 
        value={prompt} 
        onChange={(e) => setPrompt(e.target.value)} 
        placeholder="Enter natural language prompt..." 
      />
      <br/>
      <button onClick={handleSubmit}>Convert & Insert</button>
      <p>{status}</p>
    </div>
  );
}

