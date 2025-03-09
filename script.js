let clauses = {}; // Store clauses from the Excel file
let placeholders = new Set(); // Store unique placeholders
let cssFile = null; // Store the uploaded CSS file

// Load Excel file and extract clauses
document.getElementById('excelFile').addEventListener('change', function (event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet);

    console.log("Excel file parsed successfully:", json); // Debugging

    // Convert JSON to clauses dictionary
    clauses = {};
    json.forEach(row => {
      console.log("Row data:", row); // Debugging
      if (row['Clause Title'] && row['Clause Content']) {
        clauses[row['Clause Title']] = row['Clause Content'];
      } else {
        console.warn("Skipping row due to missing data:", row); // Debugging
      }
    });

    console.log("Clauses extracted:", clauses); // Debugging

    // Render clause checkboxes
    renderClauseCheckboxes();
  };
  reader.onerror = function (e) {
    console.error("Error reading file:", e); // Debugging
  };
  reader.readAsArrayBuffer(file);
});

// Handle CSS file upload
document.getElementById('cssFile').addEventListener('change', function (event) {
  const file = event.target.files[0];
  if (file) {
    cssFile = file;
    console.log("CSS file uploaded:", file.name); // Debugging
    updatePreview(); // Update preview when CSS file is uploaded
  }
});

// Render clause checkboxes
function renderClauseCheckboxes() {
  const container = document.getElementById('clauseCheckboxes');
  container.innerHTML = ''; // Clear existing checkboxes

  console.log("Rendering clauses:", Object.keys(clauses)); // Debugging

  Object.keys(clauses).forEach(clause => {
    const div = document.createElement('div');
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = clause;
    checkbox.checked = true;
    checkbox.addEventListener('change', () => {
      updatePlaceholderInputs(); // Update placeholder inputs when checkbox state changes
      updatePreview(); // Update preview in real time
    });

    const label = document.createElement('label');
    label.htmlFor = clause;
    label.textContent = clause;

    div.appendChild(checkbox);
    div.appendChild(label);
    container.appendChild(div);
  });

  updatePlaceholderInputs(); // Render placeholder inputs after checkboxes
}

// Update placeholder inputs based on selected clauses
function updatePlaceholderInputs() {
  placeholders.clear(); // Clear existing placeholders
  const container = document.getElementById('placeholderInputs');
  container.innerHTML = ''; // Clear existing inputs

  // Collect placeholders from selected clauses
  Object.keys(clauses).forEach(clause => {
    const checkbox = document.getElementById(clause);
    if (checkbox && checkbox.checked) {
      const matches = clauses[clause].match(/\{\{(.*?)\}\}/g) || [];
      matches.forEach(match => {
        const placeholder = match.slice(2, -2); // Remove {{ and }}
        placeholders.add(placeholder); // Add to the set of unique placeholders
      });
    }
  });

  // Render input fields for each placeholder
  placeholders.forEach(placeholder => {
    const div = document.createElement('div');
    const label = document.createElement('label');
    label.textContent = `Enter value for {{${placeholder}}}:`;
    const input = document.createElement('input');
    input.type = 'text';
    input.id = placeholder;
    input.addEventListener('input', updatePreview); // Update preview on input change

    div.appendChild(label);
    div.appendChild(input);
    container.appendChild(div);
  });

  console.log("Placeholders rendered:", placeholders); // Debugging
}

// Generate the contract HTML
function generateContract() {
  let htmlContent = `<html><head><title>Contract</title>`;

  // Include the uploaded CSS file
  if (cssFile) {
    htmlContent += `<link rel="stylesheet" type="text/css" href="styles.css">`;
  }

  htmlContent += `</head><body><h1>Contract</h1>`;

  // Generate Table of Contents (TOC)
  htmlContent += `<h2>Table of Contents</h2><ul>`;
  let h1Counter = 0; // Counter for <h1> headers
  let h2Counter = 0; // Counter for <h2> headers
  let h3Counter = 0; // Counter for <h3> headers

  // First pass: Collect TOC entries
  const tocEntries = [];
  Object.keys(clauses).forEach(clause => {
    const checkbox = document.getElementById(clause);
    if (checkbox.checked) {
      // Remove "**OPTIONAL**" from the clause title
      const clauseTitle = clause.replace("**OPTIONAL**", "").trim();

      // Determine the header type based on the clause title
      if (clauseTitle.startsWith("<h1>")) {
        h1Counter += 1;
        h2Counter = 0; // Reset h2 counter
        h3Counter = 0; // Reset h3 counter
        const headerText = clauseTitle.slice(4); // Remove <h1>
        tocEntries.push(`<li><a href="#h1-${h1Counter}">${h1Counter}.0 ${headerText}</a></li>`);
      } else if (clauseTitle.startsWith("<h2>")) {
        h2Counter += 1;
        h3Counter = 0; // Reset h3 counter
        const headerText = clauseTitle.slice(4); // Remove <h2>
        tocEntries.push(`<li><a href="#h2-${h1Counter}-${h2Counter}">${h1Counter}.${h2Counter} ${headerText}</a></li>`);
      } else if (clauseTitle.startsWith("<h3>")) {
        h3Counter += 1;
        const headerText = clauseTitle.slice(4); // Remove <h3>
        tocEntries.push(`<li><a href="#h3-${h1Counter}-${h2Counter}-${h3Counter}">${h1Counter}.${h2Counter}.${h3Counter} ${headerText}</a></li>`);
      }
    }
  });

  // Add TOC entries to the HTML content
  htmlContent += tocEntries.join("");
  htmlContent += `</ul>`;

  // Second pass: Generate the main content
  h1Counter = 0;
  h2Counter = 0;
  h3Counter = 0;

  Object.keys(clauses).forEach(clause => {
    const checkbox = document.getElementById(clause);
    if (checkbox.checked) {
      let content = clauses[clause];
      placeholders.forEach(placeholder => {
        const input = document.getElementById(placeholder);
        if (input) {
          content = content.replace(new RegExp(`\\{\\{${placeholder}\\}\\}`, 'g'), input.value);
        }
      });

      // Remove "**OPTIONAL**" from the clause title and content
      const clauseTitle = clause.replace("**OPTIONAL**", "").trim();
      content = content.replace("**OPTIONAL**", "").trim();

      // Determine the header type based on the clause title
      if (clauseTitle.startsWith("<h1>")) {
        h1Counter += 1;
        h2Counter = 0; // Reset h2 counter
        h3Counter = 0; // Reset h3 counter
        const headerText = clauseTitle.slice(4); // Remove <h1>
        htmlContent += `<h1 id="h1-${h1Counter}">${h1Counter}.0 ${headerText}</h1><p>${content}</p>`;
      } else if (clauseTitle.startsWith("<h2>")) {
        h2Counter += 1;
        h3Counter = 0; // Reset h3 counter
        const headerText = clauseTitle.slice(4); // Remove <h2>
        htmlContent += `<h2 id="h2-${h1Counter}-${h2Counter}">${h1Counter}.${h2Counter} ${headerText}</h2><p>${content}</p>`;
      } else if (clauseTitle.startsWith("<h3>")) {
        h3Counter += 1;
        const headerText = clauseTitle.slice(4); // Remove <h3>
        htmlContent += `<h3 id="h3-${h1Counter}-${h2Counter}-${h3Counter}">${h1Counter}.${h2Counter}.${h3Counter} ${headerText}</h3><p>${content}</p>`;
      }
    }
  });

  htmlContent += `</body></html>`;
  return htmlContent;
}

// Update the preview in real time
function updatePreview() {
  const htmlContent = generateContract();
  const previewFrame = document.getElementById('previewFrame');
  previewFrame.srcdoc = htmlContent;
}

// Download the contract and CSS file as a zip
document.getElementById('downloadButton').addEventListener('click', function () {
  const htmlContent = generateContract();
  const zip = new JSZip();

  // Add the HTML file to the zip
  zip.file("contract.html", htmlContent);

  // Add the CSS file to the zip (if uploaded)
  if (cssFile) {
    zip.file("styles.css", cssFile);
  }

  // Generate the zip file and trigger download
  zip.generateAsync({ type: "blob" })
    .then(function (content) {
      saveAs(content, "contract.zip");
    });
});

// Initial preview update
updatePreview();
