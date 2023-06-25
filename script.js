// Excel functions
const DEFAULT_SHEET_NAME = "NUEVO PROGRAMA";
const TABLE_FIELDS = [
  "Grupos",
  // "ORDEN",
  "OT",
  "PROYECTO",
  // "CRISTAL ESPECIAL",
  "CLIENTE",
  "TP\r\nOriginal\r\n",
  "Termos programados",
  "Termos fabricados",
  // "MINUTOS",
  // "M2",
  // "FECHA INGRESO",
  // "FECHA FABRICACION REAL",
  // "FECHA FABRICACION PLANIFICADA",
  // "FECHA DESPACHO",
  // "MTL Programado",
  // "MTL fabricados",
  // "MTL dimensionados fabricados",
  "Pendiente",
  // "Días Atraso",
  // "Año-mes Fab",
  // "Año-mes Plan",
  // "Cumple fecha la OT",
  // "Saldo TP",
  // "Saldo Dim",
  // "Saldo ml TP",
  // "Saldo ml Dim",
  // "URGENCIA ORM",
  // "URGENCIA SRU",
  // "URGENCIA RBP1",
  // "URGENCIA RBP2",
  // "URGENCIA",
  // "Días entre F/plan y F/Entre",
  // "m2 TP",
  // "m2 DIM",
  // "ORM",
  // "BST",
  // "SRU",
  // "Duplicado",
  // "Vendedor",
  // "Cuenta Ots",
  // "PERIODO_2",
  // "Año_Fab",
];

const fileInput = document.getElementById("inputFile");
const inputForm = document.getElementById("inputForm");

const dataTable = document.getElementById("dataTable");

async function handleFileAsync(file) {
  const data = await file.arrayBuffer();

  console.log("reading file...");
  const workbook = XLSX.read(data);

  console.log("converting data to json...");
  const worksheet = XLSX.utils.sheet_to_json(
    workbook.Sheets[DEFAULT_SHEET_NAME]
  );

  // worksheet.reverse();

  console.log("creating headers...");
  const EXCEL_HEADERS = Object.keys(worksheet[0]);
  console.log(EXCEL_HEADERS);

  TABLE_FIELDS.forEach((header) => {
    const HtmlHeader = document.createElement("th");
    HtmlHeader.textContent = header;
    dataTable.appendChild(HtmlHeader);
  });

  console.log("appending rows...");
  worksheet.forEach((row) => {
    const HtmlRow = document.createElement("tr");
    TABLE_FIELDS.forEach((field) => {
      const HtmlData = document.createElement("td");
      HtmlData.textContent = row[field];
      HtmlRow.appendChild(HtmlData);
    });
    dataTable.appendChild(HtmlRow);
  });
}

inputForm.addEventListener("submit", (event) => {
  event.preventDefault();
  Array.from(dataTable.children).forEach((e) => e.remove());
  console.log("event trigered...");
  theWorkbook = handleFileAsync(fileInput.files[0]);
});

// Filtering functions

const textQuery = (text, query) => {
  let norm_text = text.toLowerCase();
  let norm_query = query.toLowerCase();
  console.log(`Text: ${norm_text} Query: ${norm_query}`);

  return norm_text.includes(norm_query);
};

const hideHtmlElement = (element, test) => {
  if (test(element)) {
    element.hidden = !element.hidden;
  }
};

const textInclusion = (text, query) => {
  return text.toLowerCase().includes(query.toLowerCase());
};

const searchForm = document.getElementById("searchForm");
const obraInput = document.getElementById("obraInput");
const clienteInput = document.getElementById("clienteInput");

searchForm.addEventListener("submit", (event) => {
  event.preventDefault();
  let query_obra = obraInput.value;
  let query_cliente = clienteInput.value;

  const rows = Array.from(document.querySelectorAll("tr"));
  rows.forEach((row) => (row.style.display = "none"));
  rows.forEach((row) => {
    let obraMatch = textInclusion(row.innerText, query_obra);
    let clienteMatch = textInclusion(row.innerText, query_cliente);

    if (obraMatch && clienteMatch) {
      row.style.display = "";
    }
  });
});
