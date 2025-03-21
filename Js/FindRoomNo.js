async function fetchAndSearchExcel(courseCode, section) {
  try {
    const response = await fetch("../ExcelFiles/revised_final_exam_schedule_sose_undergrad_243_.xlsx"); // Adjust path
    if (!response.ok) throw new Error("File not found");

    const blob = await response.blob();
    const rows = await readXlsxFile(blob);

    if (rows.length === 0) {
      console.error("Excel file is empty");
      return;
    }

    // Identify column indexes for Course Code & Section
    const headerRow = rows[0];
    const courseCodeIndex = headerRow.indexOf("Course Code");
    const sectionIndex = headerRow.indexOf("Section");

    if (courseCodeIndex === -1 || sectionIndex === -1) {
      alert("Course Code or Section column not found!");
      return;
    }

    // Normalize input (trim spaces & convert to lowercase)
    courseCode = courseCode.trim().toLowerCase();
    section = section.trim().toLowerCase();

    // Search for matching row
    const matchedRow = rows.find((row, index) => {
      if (index === 0) return false; // Skip header row

      const courseCell = row[courseCodeIndex]?.toString().trim().toLowerCase();
      const sectionCell = row[sectionIndex]?.toString().trim().toLowerCase();

      // Check if the course code exists as a substring (case-insensitive)
      const isCourseMatch = courseCell.includes(courseCode);
      const isSectionMatch = sectionCell === section; // Exact match for section

      return isCourseMatch && isSectionMatch;
    });

    const resultCard = document.getElementById("result-card");
    resultCard.innerHTML = ""; // Clear previous results
    resultCard.classList.add("hidden");

    if (!matchedRow) {
      resultCard.innerHTML = "<p>No matching data found.</p>";
      resultCard.classList.remove("hidden");
      return;
    }

    // Create a structured result card
    matchedRow.forEach((cell, index) => {
      const item = document.createElement("div");
      item.classList.add("result-item");
      item.innerHTML = `<strong>${headerRow[index]}:</strong> ${cell ?? ""}`;
      resultCard.appendChild(item);
    });

    resultCard.classList.remove("hidden"); // Show card
    
  } catch (error) {
    console.error("Error reading file:", error);
  }
}

// Function to handle user search
function handleSearch() {
  const courseCode = document.getElementById("courseCode").value.trim();
  const section = document.getElementById("section").value.trim();

  if (courseCode && section) {
    fetchAndSearchExcel(courseCode, section);
  } else {
    alert("Please enter both Course Code and Section.");
  }
}
