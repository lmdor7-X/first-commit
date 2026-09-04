import * as XLSX from "xlsx";

import { useState } from "react";

function UploadSection() {
  const [upload1, setUpload1] = useState(null);
  const [upload2, setUpload2] = useState(null);
  const [parsedFile1, setParsedFile1] = useState([]);
  const [parsedFile2, setParsedFile2] = useState([]);
  const [result, setResult] = useState([]);
  const [filter, setFilter] = useState("All");

  const filteredResult = result.filter((item) => {
    if (filter === "All") {
      return true;
    }
    return item.Status === filter;
  });

  async function parseExcel(file) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet);

    return rows;
  }

  async function handleChange(type, e) {
    const select = e.target.files[0];

    if (type === "file1") {
      setUpload1(select);
      const rows = await parseExcel(select);
      setParsedFile1(rows);
    } else if (type === "file2") {
      setUpload2(select);
      const rows = await parseExcel(select);
      setParsedFile2(rows);
    }
  }
  const newResult = [];

  function matching() {
    const file1 = new Map();
    const file2 = new Map();
    let allMatched = true;

    parsedFile1.forEach((rows) => {
      file1.set(rows.StudentName, rows);
    });
    parsedFile2.forEach((rows) => {
      file2.set(rows.StudentName, rows);
    });
    const allStudents = new Set([...file1.keys(), ...file2.keys()]);

    const keys = Object.keys(parsedFile1[0]);
    allStudents.forEach((name) => {
      const student1 = file1.get(name);
      const student2 = file2.get(name);
      if (!student1 || !student2) {
        if (!student1) {
          newResult.push({
            Name: name,
            Status: "Not Found",
            Missingin: "File1",
            Mismatches: [],
          });
        } else {
          newResult.push({
            Name: name,
            Status: "Not Found",
            Missingin: "File2",
            Mismatches: [],
          });
        }
      } else {
        let misMatchedKeys = [];
        keys.forEach((key) => {
          if (student1[key] !== student2[key]) {
            allMatched = false;
            misMatchedKeys.push({
              Subject: key,
              File1: student1[key],
              File2: student2[key],
            });
          }
        });
        if (allMatched) {
          newResult.push({
            Name: name,
            Status: "Matched",
            Mismatches: [],
          });
        } else {
          newResult.push({
            Name: name,
            Status: "Mismatched",
            Mismatches: misMatchedKeys,
          });
          allMatched = true;
        }
      }
    });
    setResult(newResult);
  }

  const handleExport = () => {
    const exportResult = filteredResult.map((item) => ({
      Name: item.Name,
      Status: item.Status,
      Missingin:item.Missingin,
      Mismatches: item.Mismatches.map(
        (mismatch) => `${mismatch.Subject}:${mismatch.File1},${mismatch.File2}`,
      ).join(","),
    }));
    const worksheet = XLSX.utils.json_to_sheet(exportResult);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Result");
    XLSX.writeFile(workbook, "comparison_result.xlsx");
  };
  return (
    <div>
      <div id="up">
        <div>
          <h3>Upload file 1</h3>
          <input
            type="file"
            onChange={(e) => {
              handleChange("file1", e);
            }}
          />
        </div>
        <div>
          <h3>Upload File 2</h3>
          <input
            type="file"
            onChange={(e) => {
              handleChange("file2", e);
            }}
          />
        </div>
      </div>
      <button onClick={matching}>Compare</button>
      {result.length > 0 && (
        <div>
          <select value={filter} onChange={(e) => setFilter(e.target.value)}>
            <option value="All">All</option>
            <option value="Matched">Matched</option>
            <option value="Mismatched">Mismatched</option>
            <option value="Not Found">Not found</option>
          </select>
          <table border={1}>
            <thead>
              <tr>
                <th>Name</th>
                <th>Status</th>
                <th>Missingin</th>
                <th>Mismatches</th>
              </tr>
            </thead>
            <tbody>
              {filteredResult.map((item) => {
                return (
                  <tr key={item.Name}>
                    <td>{item.Name}</td>
                    <td>{item.Status}</td>
                    <td>{item.Missingin}</td>
                    <td>
                      {item.Mismatches.map((misDetail) => {
                        return (
                          <div>
                            <p>
                              {misDetail.Subject}:{misDetail.File1},
                              {misDetail.File2}
                            </p>
                          </div>
                        );
                      })}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          <button onClick={handleExport}>Download</button>
        </div>
      )}
      {result.length > 0 && (
        <div>
          Total Students={result.length} <br />
          Matched=
          {
            result.filter((item) => {
              return item.Status === "Matched";
            }).length
          }
          <br />
          Mismatched=
          {
            result.filter((item) => {
              return item.Status === "Mismatched";
            }).length
          }
          <br />
          Not Found=
          {
            result.filter((item) => {
              return item.Status === "Not Found";
            }).length
          }
          <br />
        </div>
      )}
    </div>
  );
}
export default UploadSection;
