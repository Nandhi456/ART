import { motion } from "framer-motion";
import { useState, useEffect } from "react";
import * as XLSX from "xlsx"; 
import { X, Download } from "lucide-react"; 
import { getPreview, searchPreview, openRecentFile } from "./previewApi";


const overlayStyle = {
  position: "fixed",
  inset: 0,
  background: "rgba(2, 6, 23, 0.6)",
  backdropFilter: "blur(3px)",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  padding: "24px",
  zIndex: 100,
}

const panelStyle = {
  width: "min(1920px, 100%)",
  maxHeight: "min(1080px, 90vh)",
  background: "#0f172a",
  border: "1px solid rgba(148, 163, 184, 0.2)",
  borderRadius: "16px",
  boxShadow: "0 24px 60px rgba(2, 6, 23, 0.6)",
  display: "flex",
  flexDirection: "column",
  overflow: "hidden",
}

const panelHeader = {
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  padding: "18px 22px",
  borderBottom: "1px solid rgba(148, 163, 184, 0.16)",
}

const panelTitle = {
  margin: 0,
  fontSize: "1.05rem",
  fontWeight: 700,
  color: "#f8fafc",
}

const panelSub = {
  margin: "2px 0 0",
  fontSize: "0.78rem",
  color: "#94a3b8",
}

const closeBtn = {
  border: "none",
  background: "transparent",
  color: "#94a3b8",
  cursor: "pointer",
  padding: "6px",
  borderRadius: "8px",
  display: "flex",
}

const keywordBar = {
  display: "flex",
  gap: "10px",
  alignItems: "center",
  padding: "14px 22px",
  borderBottom: "1px solid rgba(148, 163, 184, 0.16)",
  background: "rgba(17, 24, 39, 0.5)",
}

const keywordInputStyle = {
  flex: 1,
  padding: "9px 12px",
  borderRadius: "10px",
  border: "1px solid #334155",
  background: "#0f172a",
  color: "#e2e8f0",
  fontSize: "0.82rem",
  outline: "none",
}

const applyBtn = {
  border: "1px solid #334155",
  background: "transparent",
  color: "#cbd5e1",
  borderRadius: "10px",
  padding: "9px 14px",
  fontSize: "0.8rem",
  fontWeight: 600,
  cursor: "pointer",
  flexShrink: 0,
}

const exportBtn = {
  display: "flex",
  alignItems: "center",
  gap: "6px",
  border: "none",
  background: "#38bdf8",
  color: "#0f172a",
  borderRadius: "10px",
  padding: "7px 14px",
  fontSize: "0.78rem",
  fontWeight: 700,
  cursor: "pointer",
}

const scrollArea = {
  overflow: "auto",
  flex: 1,
}

const table = {
  width: "100%",
  borderCollapse: "collapse",
  fontSize: "0.82rem",
}

const th = {
  position: "sticky",
  top: 0,
  textAlign: "left",
  padding: "10px 16px",
  background: "#111827",
  color: "#7dd3fc",
  fontSize: "0.72rem",
  letterSpacing: "0.04em",
  textTransform: "uppercase",
  borderBottom: "1px solid rgba(148, 163, 184, 0.16)",
}

const td = {
  padding: "10px 16px",
  color: "#e2e8f0",
  borderBottom: "1px solid rgba(148, 163, 184, 0.08)",
}

const cellInput = {
  width: "100%",
  border: "none",
  background: "transparent",
  color: "#e2e8f0",
  fontSize: "0.82rem",
  outline: "none",
  padding: "2px 0",
}

const chip = (ok) => ({
  fontSize: "0.68rem",
  fontWeight: 700,
  padding: "3px 9px",
  borderRadius: "999px",
  background: ok ? "rgba(34, 197, 94, 0.15)" : "rgba(148, 163, 184, 0.15)",
  color: ok ? "#22c55e" : "#94a3b8",
})

const panelFooter = {
  padding: "12px 22px",
  borderTop: "1px solid rgba(148, 163, 184, 0.16)",
  fontSize: "0.78rem",
  color: "#94a3b8",
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
}


//const DEFAULT_HEADERS = ["Name", "File Name", "Qualification", "Skills", "Age", "Contact"];


/*const DUMMY_ROWS = [
  { "Name": "Vignesh", "File Name": "sample1.pdf", "Qualification": "B.Tech Mech", "Skills": "AutoCAD", "Age" : "28", "Contact" : "vignesh14@gmail.com"},
  { "Name": "Shankar", "File Name": "sample2.docx", "Qualification": "B.Tech IT", "Skills": "Java, C++, Python", "Age":"27", "Contact" : "sankar8@gmail.com"},
  { "Name": "Raja", "File Name": "sample3.pdf", "Qualification": "B.Tech ECE", "Skills": "Embedded, c++", "Age":"24", "Contact" : "raja7@gmail.com" },
  { "Name": "Suresh", "File Name": "sample4.pdf", "Qualification": "B.E CSE", "Skills": "React js, Python", "Age" : "21", "Contact" : "suresh6@gmail.com"},
  { "Name": "Raj", "File Name": "sample5.pdf", "Qualification": "B.E IT", "Skills": "Python, c++", "Age" : "22", "Contact" : "raj4@gmail.com"},
{ "Name": "Karthik", "File Name": "sample6.pdf", "Qualification": "B.Tech CSE", "Skills": "Java, Spring Boot", "Age": "26", "Contact": "karthik12@gmail.com" },
{ "Name": "Priya", "File Name": "sample7.docx", "Qualification": "B.E ECE", "Skills": "VLSI, Verilog", "Age": "23", "Contact": "priya22@gmail.com" },
{ "Name": "Arun", "File Name": "sample8.pdf", "Qualification": "B.Tech Mech", "Skills": "SolidWorks, AutoCAD", "Age": "29", "Contact": "arun9@gmail.com" },
{ "Name": "Dhivya", "File Name": "sample9.pdf", "Qualification": "B.Tech IT", "Skills": "Python, Django", "Age": "30", "Contact": "divya15@gmail.com" },
{ "Name": "Manoj", "File Name": "sample10.docx", "Qualification": "B.E CSE", "Skills": "C++, Data Structures", "Age": "25", "Contact": "manoj3@gmail.com" },
{ "Name": "Lakshmi", "File Name": "sample11.pdf", "Qualification": "B.Tech EEE", "Skills": "MATLAB, PLC", "Age": "27", "Contact": "lakshmi19@gmail.com" },
{ "Name": "Vijay", "File Name": "sample12.pdf", "Qualification": "B.E Mech", "Skills": "CATIA, ANSYS", "Age": "28", "Contact": "vijay11@gmail.com" },
{ "Name": "Sneha", "File Name": "sample13.docx", "Qualification": "B.Tech CSE", "Skills": "React js, Node js", "Age": "22", "Contact": "sneha17@gmail.com" },
{ "Name": "Prakash", "File Name": "sample14.pdf", "Qualification": "B.Tech ECE", "Skills": "Embedded C, IoT", "Age": "26", "Contact": "prakash5@gmail.com" },
{ "Name": "Anitha", "File Name": "sample15.pdf", "Qualification": "B.E IT", "Skills": "Java, SQL", "Age": "23", "Contact": "anitha21@gmail.com" },
{ "Name": "Gokul", "File Name": "sample16.docx", "Qualification": "B.Tech Mech", "Skills": "AutoCAD, CATIA", "Age": "24", "Contact": "gokul8@gmail.com" },
{ "Name": "Deepa", "File Name": "sample17.pdf", "Qualification": "B.E CSE", "Skills": "Python, Machine Learning", "Age": "25", "Contact": "deepa13@gmail.com" },
{ "Name": "Ramesh", "File Name": "sample18.pdf", "Qualification": "B.Tech EEE", "Skills": "PLC, SCADA", "Age": "29", "Contact": "ramesh16@gmail.com" },
{ "Name": "Kavya", "File Name": "sample19.docx", "Qualification": "B.Tech IT", "Skills": "Angular, JavaScript", "Age": "22", "Contact": "kavya10@gmail.com" },
{ "Name": "Senthil", "File Name": "sample20.pdf", "Qualification": "B.E ECE", "Skills": "VLSI, Embedded", "Age": "27", "Contact": "senthil6@gmail.com" },
{ "Name": "Meena", "File Name": "sample21.pdf", "Qualification": "B.Tech CSE", "Skills": "C++, Python", "Age": "23", "Contact": "meena14@gmail.com" },
{ "Name": "Arjun", "File Name": "sample22.docx", "Qualification": "B.E Mech", "Skills": "SolidWorks, ANSYS", "Age": "28", "Contact": "arjun2@gmail.com" },
{ "Name": "Nithya", "File Name": "sample23.pdf", "Qualification": "B.Tech IT", "Skills": "Java, React js", "Age": "24", "Contact": "nithya18@gmail.com" },
{ "Name": "Bala", "File Name": "sample24.pdf", "Qualification": "B.Tech ECE", "Skills": "Embedded, IoT", "Age": "26", "Contact": "bala7@gmail.com" },
{ "Name": "Swathi", "File Name": "sample25.docx", "Qualification": "B.E CSE", "Skills": "Python, Django", "Age": "21", "Contact": "swathi20@gmail.com" },
{ "Name": "Mohan", "File Name": "sample26.pdf", "Qualification": "B.Tech Mech", "Skills": "AutoCAD, SolidWorks", "Age": "29", "Contact": "mohan9@gmail.com" },
{ "Name": "Revathi", "File Name": "sample27.pdf", "Qualification": "B.Tech EEE", "Skills": "MATLAB, Simulink", "Age": "25", "Contact": "revathi4@gmail.com" },
{ "Name": "Dinesh", "File Name": "sample28.docx", "Qualification": "B.E IT", "Skills": "Java, Spring Boot", "Age": "23", "Contact": "dinesh12@gmail.com" },
{ "Name": "Pooja", "File Name": "sample29.pdf", "Qualification": "B.Tech CSE", "Skills": "React js, Node js", "Age": "22", "Contact": "pooja15@gmail.com" },
{ "Name": "Saravanan", "File Name": "sample30.pdf", "Qualification": "B.E ECE", "Skills": "Embedded C, VLSI", "Age": "27", "Contact": "saravanan3@gmail.com" },
];*/ //Dummy rows



const listVariants = {
  hidden: {},
  visible: { transition: { staggerChildren: 0.04 } },
};
const rowVariants = {
  hidden: { opacity: 0, x: -8 },
  visible: { opacity: 1, x: 0 },
};

export default function Preview({ folderName, previewData, onClose }) {
  const [rows, setRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [keywordInput, setKeywordInput] = useState("");

  useEffect(() => {
    if (previewData) {
        setRows(previewData.rows || []);
        setHeaders(previewData.columns || []);
    }
}, [previewData]);

  const applyKeywords = async () => {
    if (!keywordInput.trim()) {
        setRows(previewData.rows || []);
        return;
    }

    try {
        const result = await searchPreview(folderName, keywordInput);

        setRows(result);

    } catch (err) {
        console.error(err);
    }
};



  const buildRows=(headers)=> {
  return rows.map((known) => {
    const row = {};
    headers.forEach((h) => {
      row[h] = known[h] !== undefined ? known[h] : "";
    });
    return row;
  });
}
  

  if (!folderName) return null;

  const updateCell = (rowIndex, headerKey, value) => {
    setRows((prev) =>
      prev.map((r, i) => (i === rowIndex ? { ...r, [headerKey]: value } : r))
    );
  };

  const handleExport = () => {
    const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, `${(folderName || "export").replace(/[\\/:*?"<>|]/g, "_")}.xlsx`);
  };

  

  return (
    <motion.div
      style={overlayStyle}
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      onMouseDown={onClose}
    >
      <motion.div
        style={panelStyle}
        initial={{ opacity: 0, y: 24, scale: 0.97 }}
        animate={{ opacity: 1, y: 0, scale: 1 }}
        exit={{ opacity: 0, y: 16, scale: 0.97 }}
        transition={{ duration: 0.22, ease: [0.2, 0.8, 0.2, 1] }}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div style={panelHeader}>
          <div>
            <h3 style={panelTitle}>{folderName}</h3>
            <p style={panelSub}>{previewData?.total_files} files scanned</p>
          </div>
          <button style={closeBtn} onClick={onClose} aria-label="Close preview">
            <X size={18} />
          </button>
        </div>

        <div style={keywordBar}>
          <input
            style={keywordInputStyle}
            value={keywordInput}
            onChange={(e) => setKeywordInput(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && applyKeywords()}
            placeholder="name, email, skills, qualification"
          />
          <button style={applyBtn} onClick={applyKeywords}>
            Apply
          </button>
          <button style={exportBtn} onClick={handleExport}>
            <Download size={14} />
            Export to Excel
          </button>
        </div>

        <div style={scrollArea}>
          <table style={table}>
            <thead>
              <tr>
                {headers.map((h) => (
                  <th style={th} key={h}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <motion.tbody initial="hidden" animate="visible" variants={listVariants}>
              {rows.map((row, i) => (
                <motion.tr key={i} variants={rowVariants}>
                  {headers.map((h) => (
                    <td style={td} key={h}>
                      {h === "Status" ? (
                        <span style={chip(row[h] === "Parsed")}>{row[h] || "—"}</span>
                      ) : (
                        <input
                          style={cellInput}
                          value={row[h]}
                          onChange={(e) => updateCell(i, h, e.target.value)}
                        />
                      )}
                    </td>
                  ))}
                </motion.tr>
              ))}
            </motion.tbody>
          </table>
        </div>

        <div style={panelFooter}>
          <span>{rows.length} rows</span>
          <span>{headers.length} columns</span>
        </div>
      </motion.div>
    </motion.div>
  );
}