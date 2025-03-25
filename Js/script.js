async function fetchAndReadExcel(columnName) {
    try {
      const response = await fetch("../ExcelFiles/revised_final_exam_schedule_sose_undergrad_243_.xlsx"); // Adjust path as needed
      if (!response.ok) throw new Error("File not found");
  
      const blob = await response.blob();
      const rows = await readXlsxFile(blob);
  
      if (rows.length === 0) {
        console.error("Excel file is empty");
        return;
      }
  
      // Find the column index based on the given column name
      const headerRow = rows[0];
      const columnIndex = headerRow.indexOf(columnName);
      if (columnIndex === -1) {
        alert(`Column "${columnName}" not found`);
        return;
      }
  
      const table = document.getElementById("table-id");
      table.innerHTML = ""; // Clear previous content
  
      // Create a header row
      const headerTr = document.createElement("tr");
      const headerTd = document.createElement("th");
      headerTd.textContent = columnName;
      headerTr.appendChild(headerTd);
      table.appendChild(headerTr);
  
      // Display only the selected column
      rows.slice(1).forEach((row) => { // Skip the header row
        const tr = document.createElement("tr");
        const td = document.createElement("td");
        td.textContent = row[columnIndex] ?? ""; // Handle missing values
        tr.appendChild(td);
        table.appendChild(tr);
      });
    } catch (error) {
      console.error("Error reading file:", error);
    }
  }
  
  // Function to handle user input
  function handleFetch() {
    const columnName = document.getElementById("columnName").value.trim();
    if (columnName) {
      fetchAndReadExcel(columnName);
    } else {
      alert("Please enter a column name");
    }
  }

