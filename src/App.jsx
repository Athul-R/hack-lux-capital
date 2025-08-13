import React, { useState } from "react";
import FileUploader from "./components/FileUploader";
import SheetGrid from "./components/SheetGrid";
import VisualBasicEditor from "./components/VisualBasicEditor";
import TypeAhead from "./components/TypeAhead";
import TextPrompt from "./components/TextPrompt";
import { detectColumnTypes } from "./components/excelService";
import { ModalAIProvider } from "./components/modalAIService";

export default function App() {
  const [sheetData, setSheetData] = useState([]);
  const [columnTypes, setColumnTypes] = useState([]);
  const [vbCode, setVbCode] = useState("");

  // Called when Excel file is uploaded
  const onFileUpload = (data) => {
    setSheetData(data);
    setColumnTypes(detectColumnTypes(data));
  };

  return (
    <ModalAIProvider>
      <div className="app-container">
        <h1>Browser Excel & Visual Basic Editor</h1>
        <FileUploader onFileUpload={onFileUpload} />
        <SheetGrid sheetData={sheetData} columnTypes={columnTypes} />
        <VisualBasicEditor vbCode={vbCode} setVbCode={setVbCode} />
        <TypeAhead sheetData={sheetData} />
        <TextPrompt sheetData={sheetData} setSheetData={setSheetData} />
      </div>
    </ModalAIProvider>
  );
}
