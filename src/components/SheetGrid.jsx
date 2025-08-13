import React from "react";

export default function SheetGrid({ sheetData, columnTypes }) {
  if (!sheetData || sheetData.length === 0) {
    return <div>No data loaded.</div>;
  }

  return (
    <table>
      <thead>
        <tr>
          {sheetData[0].map((colName, idx) => (
            <th key={idx}>
              {colName} <small>({columnTypes[idx] || "string"})</small>
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {sheetData.slice(1).map((row, ridx) => (
          <tr key={ridx}>
            {row.map((cell, cidx) => (
              <td key={cidx}>{cell}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
