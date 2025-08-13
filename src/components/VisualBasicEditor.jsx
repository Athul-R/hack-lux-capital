import React from "react";
import MonacoEditor from "react-monaco-editor";

export default function VisualBasicEditor({ vbCode, setVbCode }) {
  const options = {
    selectOnLineNumbers: true,
    minimap: { enabled: false }
  };

  return (
    <div className="vb-editor">
      <h3>Visual Basic Editor</h3>
      <MonacoEditor
        language="vb"
        value={vbCode}
        options={options}
        onChange={setVbCode}
        height="180"
      />
    </div>
  );
}
