// Variables globales
let dataMatrix = [];
let headers = ["LAST", "CURRENT", "NEXT", "NEXT +1"];
let displayData = [];
let showNetChangesOnly = false;

// Variables para saldos por período (globales)
let saldoBeginningCash = [0, 0, 0, 0];
let saldoChangeCashEquiv = [0, 0, 0, 0];
let saldoEndingCash = [0, 0, 0, 0];

// Función para formatear números en $ con separador de miles y decimal
function formatCurrency(value) {
  if (isNaN(value) || value === null || value === undefined) return "";
  return "$" + value.toLocaleString("es-CL", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Función para crear y descargar el template en XLSX
function downloadTemplate() {
  const ws_data = [];

  // Encabezado principal
  ws_data.push(["ACCOUNT", ...headers]);

  // Secciones y items
  const sections = [
    {
      title: "CASH FLOW FROM OPERATING ACTIVITIES",
      items: [
        { name: "Net Income", values: ["", "", "", ""] },
        { name: "Adjustments to reconcile net income to net cash provided by operating activities", values: ["", "", "", ""] },
        { name: "Depreciation on fixed assets", values: ["", "", "", ""] },
        { name: "GoodWill / Intangible Impairments", values: ["", "", "", ""] },
        { name: "Deferred Taxes", values: ["", "", "", ""] },
      ],
    },
    {
      title: "Increase (decrease) in current assets",
      items: [
        { name: "accounts receivable", values: ["", "", "", ""] },
        { name: "Inventory", values: ["", "", "", ""] },
        { name: "Prepaid Expenses", values: ["", "", "", ""] },
      ],
    },
    {
      title: "Increase (decrease) in current liabilities",
      items: [
        { name: "Accounts Payable", values: ["", "", "", ""] },
        { name: "Accrue expenses and unearned revenues", values: ["", "", "", ""] },
      ],
    },
    {
      title: "Net cash provided by operating activities",
      items: [
        { name: "Total Operating Cash Flow", values: ["", "", "", ""] },
      ],
    },
    {
      title: "CASH FLOW FROM INVESTING ACTIVITIES",
      items: [
        { name: "Purchase of Property and Equipment", values: ["", "", "", ""] },
        { name: "Additions to Intangibles", values: ["", "", "", ""] },
      ],
    },
    {
      title: "Net Cash Used In Investing Activities",
      items: [
        { name: "Total Investing Cash Flow", values: ["", "", "", ""] },
      ],
    },
    {
      title: "CASH FLOW FROM FINANCING ACTIVITIES",
      items: [
        { name: "Proceeds from line of credit", values: ["", "", "", ""] },
        { name: "Payments on line of credit", values: ["", "", "", ""] },
        { name: "Proceeds from long-term debt", values: ["", "", "", ""] },
        { name: "Payments on long-term debt", values: ["", "", "", ""] },
        { name: "Dividends Paid", values: ["", "", "", ""] },
        { name: "Stock Issuances / Repurchases", values: ["", "", "", ""] },
        { name: "Debt issuances / Repayments", values: ["", "", "", ""] },
      ],
    },
    {
      title: "Net Cash Provided by financing Activities",
      items: [
        { name: "Total Financing Cash Flow", values: ["", "", "", ""] },
      ],
    },
    {
      title: "Balance",
      items: [
        { name: "Beginning Cash", values: ["", "", "", ""] },
        { name: "Net Change in Cash & Cash Equivalents", values: ["", "", "", ""] },
        { name: "Ending Cash", values: ["", "", "", ""] },
      ],
    },
  ];

  // Construir matriz
  sections.forEach(section => {
    if (section.title) {
      ws_data.push([section.title, "", "", "", ""]);
    }
    section.items.forEach(item => {
      ws_data.push([item.name, ...item.values]);
    });
  });

  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Template");
  XLSX.writeFile(wb, "template_financiero.xlsx");
}

// Eventos para botones
document.getElementById("downloadTemplate").addEventListener("click", () => {
  downloadTemplate();
});
document.getElementById("loadFileBtn").addEventListener("click", () => {
  document.getElementById("uploadFile").click();
});
document.getElementById("uploadFile").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });
    dataMatrix = jsonData;
    processData();
  };
  reader.readAsArrayBuffer(file);
});

// Función para procesar y mostrar datos
function processData() {
  displayData = [];
  for (let i = 0; i < dataMatrix.length; i++) {
    const row = dataMatrix[i];
    let values = row.slice(1, 5).map(v => {
      const num = parseFloat(v);
      return isNaN(num) || v === null || v === undefined ? 0 : num;
    });
    const total = values.reduce((a, b) => a + b, 0);
    displayData.push({ row, total });
  }
  renderTable();
}

// Función para renderizar la tabla
function renderTable() {
  const container = document.getElementById("financials");
  container.innerHTML = ""; // limpiar

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  // Encabezado
  const trHead = document.createElement("tr");
  // Encabezados en la vista: ACCOUNT | LAST | CURRENT | NEXT | NEXT +1
  ["ACCOUNT", ...headers].forEach(text => {
    const th = document.createElement("th");
    th.innerText = text;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);
  table.appendChild(thead);

  // Variables para saldos por período
  let localSaldoBeginningCash = [0, 0, 0, 0];
  let localSaldoChangeCash = [0, 0, 0, 0];
  let localSaldoEndingCash = [0, 0, 0, 0];

  // Función para actualizar saldos por línea
  const updateSaldosPorLinea = (nombreLinea, valores) => {
    const lineLower = nombreLinea.toLowerCase();
    if (lineLower.includes("beginning cash")) {
      saldoBeginningCash = [...valores];
      localSaldoBeginningCash = [...valores];
    } else if (lineLower.includes("change in cash & cash")) {
      saldoChangeCashEquiv = [...valores];
      localSaldoChangeCash = [...valores];
    } else if (lineLower.includes("ending cash")) {
      saldoEndingCash = [...valores];
      localSaldoEndingCash = [...valores];
    }
  };

  // Variable para marcar si ya agregamos las filas de saldo
  let saldoFilasAgregadas = false;

  // Función para agregar fila de saldo
  const addSaldoRow = (label, valores) => {
    const trSaldo = document.createElement("tr");
    trSaldo.style.fontWeight = "bold";

    const tdLabel = document.createElement("td");
    tdLabel.innerText = label;
    trSaldo.appendChild(tdLabel);

    for (let i = 0; i < 4; i++) {
      const tdVal = document.createElement("td");
      tdVal.innerText = formatCurrency(valores[i]);
      if (valores[i] > 0) tdVal.style.color = "green";
      else if (valores[i] < 0) tdVal.style.color = "red";
      trSaldo.appendChild(tdVal);
    }
    // Rellenar columnas vacías si es necesario
    for (let i = 0; i < 4 - valores.length; i++) {
      const tdVacio = document.createElement("td");
      tdVacio.innerText = "";
      trSaldo.appendChild(tdVacio);
    }
    tbody.appendChild(trSaldo);
  };

  // Filtrar filas a mostrar
  let filasMostradas = displayData.filter(d => {
    const titulo = d.row[0]?.toString() || "";
    const tituloUpper = titulo.toUpperCase();

    if (showNetChangesOnly) {
      // Mostrar solo las líneas de net cash flows
      return (
        tituloUpper === "NET CASH PROVIDED BY OPERATING ACTIVITIES" ||
        tituloUpper === "NET CASH USED IN INVESTING ACTIVITIES" ||
        tituloUpper === "NET CASH PROVIDED BY FINANCING ACTIVITIES"
      );
    }
    // Mostrar filas con valores distintos de cero
    const valores = d.row.slice(1, 5).map(v => {
      const num = parseFloat(v);
      return isNaN(num) || v === null || v === undefined ? 0 : num;
    });
    return valores.some(v => v !== 0);
  });

  // Mostrar filas filtradas y actualizar saldos
  filasMostradas.forEach(d => {
    const titulo = d.row[0]?.toString() || "";
    const lineNameLower = titulo.toLowerCase();

    // Encabezados
    if (lineNameLower.includes("encabezado")) {
      const trHeader = document.createElement("tr");
      trHeader.style.backgroundColor = "#4B0082";
      trHeader.style.color = "#00FFCC";

      const td = document.createElement("td");
      td.innerText = titulo;
      td.colSpan = 5;
      td.style.fontWeight = 'bold';
      trHeader.appendChild(td);
      tbody.appendChild(trHeader);
      return;
    }

    // Reemplazar filas sin saldo por ceros en modo net changes
    let valores = d.row.slice(1, 5).map(v => {
      const num = parseFloat(v);
      // Si estamos en modo net changes, reemplazar null o vacío por ceros
      if (showNetChangesOnly) {
        return isNaN(num) || v === null || v === undefined || v === "" ? 0 : num;
      } else {
        return isNaN(num) || v === null || v === undefined ? 0 : num;
      }
    });

    if (showNetChangesOnly) {
      // Ver si todos los valores son 0, si es así, los reemplazamos por ceros explícitos
      const todosCeros = valores.every(v => v === 0);
      if (todosCeros) {
        valores = [0, 0, 0, 0];
      }
    }

    // Actualizar saldos si corresponde
    updateSaldosPorLinea(titulo, valores);

    // Mostrar fila
    const trFila = document.createElement("tr");
    d.row.forEach((cell, idx) => {
      const td = document.createElement("td");
      if (titulo.toUpperCase().includes("ENCABEZADO")) {
        td.style.backgroundColor = "#4B0082";
        td.style.color = "#00FFCC";
        td.innerText = cell;
      } else {
        if (idx > 0) {
          const val = parseFloat(cell);
          if (!isNaN(val) && val !== null && val !== undefined && cell !== "") {
            td.innerText = formatCurrency(val);
            if (val > 0) td.style.color = "green";
            else if (val < 0) td.style.color = "red";
          } else {
            td.innerText = "";
          }
        } else {
          // Primera columna
          td.innerText = cell;
        }
      }
      trFila.appendChild(td);
    });
    tbody.appendChild(trFila);
  });

  // Agregar filas de saldo solo en modo net changes
  if (showNetChangesOnly && !saldoFilasAgregadas) {
    // Agregar saldo de Beginning Cash
    addSaldoRow("Beginning Cash", saldoBeginningCash);
    // Agregar saldo de Change in Cash & Cash
    addSaldoRow("Change in Cash & Cash", saldoChangeCashEquiv);
    // Agregar saldo de Ending Cash
    addSaldoRow("Ending Cash", saldoEndingCash);

    saldoFilasAgregadas = true;
  }

  table.appendChild(tbody);
  container.appendChild(table);
}

// Función para toggle de "Net Changes"
function toggleNetChanges() {
  showNetChangesOnly = !showNetChangesOnly;
  renderTable();
}
document.getElementById("toggleNetChangesBtn").addEventListener("click", toggleNetChanges);
