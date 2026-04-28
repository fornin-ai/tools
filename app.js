document.getElementById('upload').addEventListener('change', handleFile);

let tableInstance;

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const json = XLSX.utils.sheet_to_json(worksheet);

    render(json);
  };

  reader.readAsArrayBuffer(file);
}

function render(data) {

  if (tableInstance) {
    tableInstance.destroy();
    $('#table').empty();
  }

  tableInstance = $('#table').DataTable({
    data: data,
    columns: [
      { title: "Control ID", data: "Control ID" },
      { title: "Control Name", data: "Control Name" },
      { title: "Category", data: "Control Category" },
      { title: "Standard", data: "Referenced Standard(s)" },
      { title: "Reference", data: "Standard Article Reference" }
    ]
  });

  // Row click for details popup
  $('#table tbody').off('click').on('click', 'tr', function () {
    const rowData = tableInstance.row(this).data();

    if (!rowData) return;

    const container = document.createElement("div");
    container.style.textAlign = "left";

    function addSection(title, content) {
      const titleEl = document.createElement("div");
      titleEl.style.fontWeight = "bold";
      titleEl.style.marginTop = "10px";
      titleEl.textContent = title;

      const contentEl = document.createElement("div");
      contentEl.style.whiteSpace = "pre-wrap";
      contentEl.textContent = content || "";

      container.appendChild(titleEl);
      container.appendChild(contentEl);
    }

    addSection("Control:", rowData["Control Name"]);
    addSection("Description:", rowData["Short Description"]);
    addSection("Implementation:", rowData["Detailed Technical Implementation"]);

    Swal.fire({
      title: "Control Details",
      html: container,
      width: 700
    });
  });
}
