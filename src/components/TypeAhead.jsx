import React, { useState } from "react";

const formulas = [
  "VLOOKUP(", "INDEX(", "MATCH(", "SUM(", "SUMIF(", 
  "COUNTIF(", "IF(", "CONCAT(", "PIVOT TABLE", "SORT("
];

export default function TypeAhead({ sheetData }) {
  const [input, setInput] = useState("");
  const [suggestions, setSuggestions] = useState([]);

  const onChange = (e) => {
    const val = e.target.value;
    setInput(val);
    if (val.length > 0) {
      const filtered = formulas.filter(f => f.toLowerCase().startsWith(val.toLowerCase()));
      setSuggestions(filtered);
    } else {
      setSuggestions([]);
    }
  };

  return (
    <div className="typeahead">
      <h3>Formula Suggestions</h3>
      <input type="text" value={input} onChange={onChange} placeholder="Start typing formula..." />
      <ul>
        {suggestions.map((s, i) => (
          <li key={i}>{s}</li>
        ))}
      </ul>
    </div>
  );
}
