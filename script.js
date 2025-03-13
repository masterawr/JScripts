let clauses = {}; // Store clauses from the Excel file
let placeholders = new Set(); // Store unique placeholders
let clauseIds = {}; // Store header IDs for each clause

// Initialize the preview iframe with a Shadow DOM
function initializePreview() {
  const previewFrame = document.getElementById('previewFrame');
  const previewDocument = previewFrame.contentDocument;

  // Create a Shadow DOM root inside the iframe's body
  const shadowRoot = previewDocument.body.attachShadow({ mode: 'open' });

  // Add a container for the preview content
  shadowRoot.innerHTML = `<div id="preview-content"></div>`;
}

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

    // Add a "Go to Clause" button
    const goToButton = document.createElement('button');
    goToButton.textContent = "Go to Clause";
    goToButton.className = "go-to-clause-button";
    goToButton.addEventListener('click', () => {
      scrollToClause(clause); // Scroll to the clause in the preview
    });

    div.appendChild(checkbox);
    div.appendChild(label);
    div.appendChild(goToButton); // Add the button next to the clause
    container.appendChild(div);
  });

  updatePlaceholderInputs(); // Render placeholder inputs after checkboxes
}

// Scroll to a specific clause in the preview
function scrollToClause(clauseTitle) {
  const previewFrame = document.getElementById('previewFrame');
  const shadowRoot = previewFrame.contentDocument.body.shadowRoot;

  // Get the header ID for the clause
  const headerId = clauseIds[clauseTitle];
  if (headerId) {
    const targetElement = shadowRoot.getElementById(headerId);
    if (targetElement) {
      targetElement.scrollIntoView({ behavior: 'smooth' });
    } else {
      console.warn("Target element not found:", headerId); // Debugging
    }
  } else {
    console.warn("Header ID not found for clause:", clauseTitle); // Debugging
  }
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

  // Add smooth scrolling CSS
  htmlContent += `
    <style>
      html {
        scroll-behavior: smooth;
      }

      /* Style for Table of Contents links */
      a {
        color: #27ae60; /* Vibrant green for links */
        text-decoration: none; /* Remove underline */
        transition: color 0.3s ease; /* Smooth transition for hover effect */
      }

      a:hover {
        color: #e67e22; /* Warm orange for hover state */
      }
    </style>
  `;

  htmlContent += `</head><body><h1>Contract</h1>`;

  // Generate Table of Contents (TOC)
  htmlContent += `<h2>Table of Contents</h2><ul>`;

  // First pass: Collect TOC entries and reset counters
  let h1Counter = 0; // Start at 0 for <h1> headers
  let h2Counter = 0; // Start at 0 for <h2> headers
  let h3Counter = 0; // Start at 0 for <h3> headers

  const tocEntries = [];
  Object.keys(clauses).forEach(clause => {
    const checkbox = document.getElementById(clause);
    if (checkbox.checked) {
      // Remove "**OPTIONAL**" from the clause title
      const clauseTitle = clause.replace("**OPTIONAL**", "").trim();

      // Determine the header type based on the clause title
      if (clauseTitle.startsWith("<h1>")) {
        h1Counter += 1; // Increment h1 counter
        h2Counter = 0; // Reset h2 counter
        h3Counter = 0; // Reset h3 counter

        const headerText = clauseTitle.slice(4); // Remove <h1>
        const headerId = `h1-${h1Counter}`;
        clauseIds[clause] = headerId; // Store the header ID for this clause
        tocEntries.push(`<li><a href="#${headerId}" class="toc-link">${h1Counter}.0 ${headerText}</a></li>`);
      } else if (clauseTitle.startsWith("<h2>")) {
        h2Counter += 1; // Increment h2 counter
        h3Counter = 0; // Reset h3 counter

        const headerText = clauseTitle.slice(4); // Remove <h2>
        const headerId = `h2-${h1Counter}-${h2Counter}`;
        clauseIds[clause] = headerId; // Store the header ID for this clause
        tocEntries.push(`<li><a href="#${headerId}" class="toc-link">${h1Counter}.${h2Counter} ${headerText}</a></li>`);
      } else if (clauseTitle.startsWith("<h3>")) {
        h3Counter += 1; // Increment h3 counter

        const headerText = clauseTitle.slice(4); // Remove <h3>
        const headerId = `h3-${h1Counter}-${h2Counter}-${h3Counter}`;
        clauseIds[clause] = headerId; // Store the header ID for this clause
        tocEntries.push(`<li><a href="#${headerId}" class="toc-link">${h1Counter}.${h2Counter}.${h3Counter} ${headerText}</a></li>`);
      }
    }
  });

  // Add TOC entries to the HTML content
  htmlContent += tocEntries.join("");
  htmlContent += `</ul>`;

  // Second pass: Generate the main content
  h1Counter = 0; // Reset h1 counter to 0
  h2Counter = 0; // Reset h2 counter to 0
  h3Counter = 0; // Reset h3 counter to 0

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
        h1Counter += 1; // Increment h1 counter
        h2Counter = 0; // Reset h2 counter
        h3Counter = 0; // Reset h3 counter

        const headerText = clauseTitle.slice(4); // Remove <h1>
        const headerId = `h1-${h1Counter}`;
        htmlContent += `<h1 id="${headerId}">${h1Counter}.0 ${headerText}</h1><p>${content}</p>`;
      } else if (clauseTitle.startsWith("<h2>")) {
        h2Counter += 1; // Increment h2 counter
        h3Counter = 0; // Reset h3 counter

        const headerText = clauseTitle.slice(4); // Remove <h2>
        const headerId = `h2-${h1Counter}-${h2Counter}`;
        htmlContent += `<h2 id="${headerId}">${h1Counter}.${h2Counter} ${headerText}</h2><p>${content}</p>`;
      } else if (clauseTitle.startsWith("<h3>")) {
        h3Counter += 1; // Increment h3 counter

        const headerText = clauseTitle.slice(4); // Remove <h3>
        const headerId = `h3-${h1Counter}-${h2Counter}-${h3Counter}`;
        htmlContent += `<h3 id="${headerId}">${h1Counter}.${h2Counter}.${h3Counter} ${headerText}</h3><p>${content}</p>`;
      }
    }
  });

  htmlContent += `</body></html>`;
  return htmlContent;
}

// Update the preview in real time
function updatePreview() {
  const previewFrame = document.getElementById('previewFrame');
  const shadowRoot = previewFrame.contentDocument.body.shadowRoot;

  // Store the current scroll position of the iframe
  const scrollPosition = previewFrame.contentWindow.scrollY;

  // Generate the new HTML content
  const htmlContent = generateContract();

  // Update the preview content inside the Shadow DOM
  shadowRoot.getElementById('preview-content').innerHTML = htmlContent;

  // Attach event listeners to TOC links
  const tocLinks = shadowRoot.querySelectorAll('.toc-link');
  tocLinks.forEach(link => {
    link.addEventListener('click', function (event) {
      event.preventDefault(); // Prevent default link behavior
      const targetId = link.getAttribute('href').slice(1); // Remove the '#' from href
      const targetElement = shadowRoot.getElementById(targetId);
      if (targetElement) {
        targetElement.scrollIntoView({ behavior: 'smooth' }); // Smooth scroll to the target element
      }
    });
  });

  // Restore the scroll position after updating the content
  previewFrame.contentWindow.scrollTo(0, scrollPosition);
}

// Initialize the preview iframe when the page loads
initializePreview();

// Download the contract as a zip
document.getElementById('downloadButton').addEventListener('click', function () {
  const htmlContent = generateContract();
  const zip = new JSZip();

  // Add the HTML file to the zip
  zip.file("contract.html", htmlContent);

  // Generate the zip file and trigger download
  zip.generateAsync({ type: "blob" })
    .then(function (content) {
      saveAs(content, "contract.zip");
    });
});

// Initial preview update
updatePreview();
