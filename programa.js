const HEADER_ALIASES = {
  "codigo interno": "codigo_interno",
  producto: "producto",
  "descripcion adicional": "descripcion_adicional",
  marca: "marca",
  indice: "indice",
  "unidad caja granel": "unidad_caja_granel",
  "unidad caja fraccion": "unidad_caja_fraccion",
  "unidad caja fracción": "unidad_caja_fraccion",
  "precio de lista unitario": "precio_lista_unitario",
  "precio de lista neto unitario": "precio_lista_neto_unitario",
  "precio de lista neto 100 unid": "precio_lista_neto_100_unid",
};

const REQUIRED_FIELDS = [
  "codigo_interno",
  "producto",
  "descripcion_adicional",
  "marca",
  "precio_lista_unitario",
  "precio_lista_neto_unitario",
  "precio_lista_neto_100_unid",
];

const PRICE_MODES = {
  "Lista unitario": "precio_lista_unitario",
  "Neto unitario": "precio_lista_neto_unitario",
  "Neto 100 unid (prorrateado)": "precio_lista_neto_100_unid",
};

const state = {
  products: [],
  filteredProducts: [],
  quoteItems: [],
  currentPriceFile: "",
};

const els = {
  priceFile: document.querySelector("#priceFile"),
  pickFile: document.querySelector("#pickFile"),
  exportQuote: document.querySelector("#exportQuote"),
  clearQuote: document.querySelector("#clearQuote"),
  statusText: document.querySelector("#statusText"),
  productCount: document.querySelector("#productCount"),
  quoteNumber: document.querySelector("#quoteNumber"),
  quoteDate: document.querySelector("#quoteDate"),
  clientName: document.querySelector("#clientName"),
  clientPhone: document.querySelector("#clientPhone"),
  clientAddress: document.querySelector("#clientAddress"),
  search: document.querySelector("#search"),
  quantity: document.querySelector("#quantity"),
  priceMode: document.querySelector("#priceMode"),
  productsBody: document.querySelector("#productsBody"),
  quoteList: document.querySelector("#quoteList"),
  emptyQuote: document.querySelector("#emptyQuote"),
  discount: document.querySelector("#discount"),
  iva: document.querySelector("#iva"),
  extra: document.querySelector("#extra"),
  subtotal: document.querySelector("#subtotal"),
  discountAmount: document.querySelector("#discountAmount"),
  ivaAmount: document.querySelector("#ivaAmount"),
  total: document.querySelector("#total"),
};

function normalizeText(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function toFloat(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;
  let text = String(value).trim().replace("$", "").replaceAll(" ", "");
  if (text.includes(",") && text.includes(".")) text = text.replaceAll(".", "").replace(",", ".");
  else if (text.includes(",")) text = text.replace(",", ".");
  const number = Number(text);
  return Number.isFinite(number) ? number : 0;
}

function positiveNumber(input) {
  const value = toFloat(input.value);
  return value >= 0 ? value : 0;
}

function money(value) {
  return new Intl.NumberFormat("es-AR", {
    style: "currency",
    currency: "ARS",
  }).format(Number(value || 0));
}

function autoQuoteNumber() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return `PRES-${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

function todayText() {
  return new Intl.DateTimeFormat("es-AR").format(new Date());
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function toast(message) {
  const node = document.createElement("div");
  node.className = "toast";
  node.textContent = message;
  document.body.appendChild(node);
  window.setTimeout(() => node.remove(), 3200);
}

function findHeaderRow(rows) {
  const maxRows = Math.min(rows.length, 40);
  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const values = (rows[rowIndex] || []).map(normalizeText);
    if (values.includes("codigo interno") && values.includes("producto")) {
      return rowIndex;
    }
  }
  throw new Error("No se encontró la fila de encabezados. El Excel debe tener columnas como CODIGO INTERNO y PRODUCTO.");
}

function buildHeaderMap(headerValues) {
  const mapping = {};
  headerValues.forEach((value, index) => {
    const key = HEADER_ALIASES[normalizeText(value)];
    if (key) mapping[key] = index;
  });

  const missing = REQUIRED_FIELDS.filter((field) => mapping[field] === undefined);
  if (missing.length) {
    throw new Error(`Faltan columnas requeridas en el Excel: ${missing.join(", ")}`);
  }
  return mapping;
}

function cell(row, index) {
  return index >= 0 && index < row.length ? row[index] : "";
}

function loadProductsFromWorkbook(workbook) {
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) throw new Error("El archivo no tiene hojas.");

  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headerRow = findHeaderRow(rows);
  const headerMap = buildHeaderMap(rows[headerRow]);
  const products = [];

  for (const row of rows.slice(headerRow + 1)) {
    const code = cell(row, headerMap.codigo_interno);
    const productName = cell(row, headerMap.producto);
    const description = cell(row, headerMap.descripcion_adicional);
    const brand = cell(row, headerMap.marca);
    if (![code, productName, description, brand].some((value) => String(value).trim())) continue;

    const item = {
      codigo_interno: String(code ?? "").trim(),
      producto: String(productName ?? "").trim(),
      descripcion_adicional: String(description ?? "").trim(),
      marca: String(brand ?? "").trim(),
      indice: String(cell(row, headerMap.indice ?? -1) ?? "").trim(),
      unidad_caja_granel: toFloat(cell(row, headerMap.unidad_caja_granel ?? -1)),
      unidad_caja_fraccion: toFloat(cell(row, headerMap.unidad_caja_fraccion ?? -1)),
      precio_lista_unitario: toFloat(cell(row, headerMap.precio_lista_unitario)),
      precio_lista_neto_unitario: toFloat(cell(row, headerMap.precio_lista_neto_unitario)),
      precio_lista_neto_100_unid: toFloat(cell(row, headerMap.precio_lista_neto_100_unid)),
    };
    item.search_text = normalizeText([
      item.codigo_interno,
      item.producto,
      item.descripcion_adicional,
      item.marca,
      item.indice,
    ].join(" "));
    products.push(item);
  }

  if (!products.length) {
    throw new Error("El archivo se abrió, pero no se encontraron productos debajo del encabezado.");
  }
  return products;
}

async function loadPriceFile(file) {
  if (!window.XLSX) {
    throw new Error("No se pudo cargar la librería de Excel. Revisá la conexión a internet.");
  }

  els.statusText.textContent = "Leyendo lista de precios...";
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  state.products = loadProductsFromWorkbook(workbook);
  state.currentPriceFile = file.name;
  filterProducts();
  els.productCount.textContent = `${state.products.length.toLocaleString("es-AR")} productos`;
  els.statusText.textContent = `Lista cargada: ${file.name}`;
  toast("Lista cargada correctamente.");
}

function filterProducts() {
  const terms = normalizeText(els.search.value).split(" ").filter(Boolean);
  state.filteredProducts = terms.length
    ? state.products.filter((product) => terms.every((term) => product.search_text.includes(term)))
    : [...state.products];
  renderProducts();
}

function renderProducts() {
  if (!state.products.length) {
    els.productsBody.innerHTML = '<tr><td colspan="7">Cargá una lista de precios para ver productos.</td></tr>';
    return;
  }

  const visibleProducts = state.filteredProducts.slice(0, 500);
  if (!visibleProducts.length) {
    els.productsBody.innerHTML = '<tr><td colspan="7">No hay productos para mostrar.</td></tr>';
    return;
  }

  els.productsBody.innerHTML = visibleProducts.map((product, index) => `
    <tr>
      <td>${escapeHtml(product.codigo_interno)}</td>
      <td>
        <strong>${escapeHtml(product.producto)}</strong>
        <div class="muted">${escapeHtml(product.descripcion_adicional)}</div>
      </td>
      <td>${escapeHtml(product.marca)}</td>
      <td class="money">${money(product.precio_lista_unitario)}</td>
      <td class="money">${money(product.precio_lista_neto_unitario)}</td>
      <td class="money">${money(product.precio_lista_neto_100_unid)}</td>
      <td><button class="row-add" type="button" data-index="${index}" title="Agregar producto">+</button></td>
    </tr>
  `).join("");
}

function addProduct(product) {
  const quantity = positiveNumber(els.quantity);
  if (quantity <= 0) {
    toast("Ingresá una cantidad mayor que cero.");
    return;
  }

  const priceMode = els.priceMode.value;
  const priceField = PRICE_MODES[priceMode];
  let unitPrice = Number(product[priceField] || 0);
  if (priceField === "precio_lista_neto_100_unid") {
    unitPrice = unitPrice ? unitPrice / 100 : 0;
  }

  const detail = [product.producto, product.descripcion_adicional, product.marca]
    .filter(Boolean)
    .join(" | ");

  state.quoteItems.push({
    codigo: product.codigo_interno,
    detalle: `${detail} (${priceMode})`,
    cantidad: quantity,
    precio_unitario: unitPrice,
  });
  renderQuote();
}

function renderQuote() {
  els.emptyQuote.classList.toggle("visible", state.quoteItems.length === 0);
  els.quoteList.innerHTML = state.quoteItems.map((item, index) => `
    <article class="quote-item">
      <div>
        <strong>${escapeHtml(item.detalle)}</strong>
        <div class="muted">${escapeHtml(item.codigo)} · ${item.cantidad.toLocaleString("es-AR")} x ${money(item.precio_unitario)}</div>
        <strong>${money(item.cantidad * item.precio_unitario)}</strong>
      </div>
      <button class="remove" type="button" data-index="${index}" title="Quitar">×</button>
    </article>
  `).join("");
  renderTotals();
}

function getTotals() {
  const subtotal = state.quoteItems.reduce((sum, item) => sum + item.cantidad * item.precio_unitario, 0);
  const discountPct = positiveNumber(els.discount);
  const ivaPct = positiveNumber(els.iva);
  const extraCost = positiveNumber(els.extra);
  const discountAmount = subtotal * discountPct / 100;
  const base = subtotal - discountAmount;
  const ivaAmount = base * ivaPct / 100;
  const total = base + ivaAmount + extraCost;
  return { subtotal, discountPct, ivaPct, extraCost, discountAmount, ivaAmount, total };
}

function renderTotals() {
  const totals = getTotals();
  els.subtotal.textContent = money(totals.subtotal);
  els.discountAmount.textContent = money(totals.discountAmount);
  els.ivaAmount.textContent = money(totals.ivaAmount);
  els.total.textContent = money(totals.total);
}

function exportQuote() {
  if (!state.quoteItems.length) {
    toast("Agregá al menos un producto antes de exportar.");
    return;
  }

  if (!window.XLSX) {
    toast("No se pudo cargar la librería de Excel.");
    return;
  }

  const totals = getTotals();
  const rows = [
    ["PRESUPUESTO"],
    [],
    ["Número", els.quoteNumber.value],
    ["Fecha", els.quoteDate.value],
    ["Cliente", els.clientName.value],
    ["Teléfono", els.clientPhone.value],
    ["Dirección / Observación", els.clientAddress.value],
    ["Lista de precios usada", state.currentPriceFile],
    [],
    ["#", "Código", "Detalle", "Cantidad", "Precio unitario", "Subtotal"],
  ];

  state.quoteItems.forEach((item, index) => {
    rows.push([
      index + 1,
      item.codigo,
      item.detalle,
      item.cantidad,
      item.precio_unitario,
      item.cantidad * item.precio_unitario,
    ]);
  });

  rows.push(
    [],
    ["", "", "", "", "Subtotal", totals.subtotal],
    ["", "", "", "", `Descuento (${totals.discountPct.toFixed(2)}%)`, totals.discountAmount],
    ["", "", "", "", `IVA (${totals.ivaPct.toFixed(2)}%)`, totals.ivaAmount],
    ["", "", "", "", "Recargo / Flete", totals.extraCost],
    ["", "", "", "", "TOTAL", totals.total],
  );

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  worksheet["!cols"] = [
    { wch: 8 },
    { wch: 16 },
    { wch: 60 },
    { wch: 12 },
    { wch: 18 },
    { wch: 18 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Presupuesto");
  const fileName = `${(els.quoteNumber.value || autoQuoteNumber()).replace(/[^\w.-]+/g, "_")}.xlsx`;
  XLSX.writeFile(workbook, fileName);
  toast("Presupuesto exportado.");
}

function clearQuote() {
  state.quoteItems = [];
  els.quoteNumber.value = autoQuoteNumber();
  els.clientName.value = "";
  els.clientPhone.value = "";
  els.clientAddress.value = "";
  els.discount.value = "0";
  els.iva.value = "21";
  els.extra.value = "0";
  renderQuote();
}

function setup() {
  els.quoteNumber.value = autoQuoteNumber();
  els.quoteDate.value = todayText();
  els.priceMode.innerHTML = Object.keys(PRICE_MODES)
    .map((mode) => `<option value="${mode}">${mode}</option>`)
    .join("");

  els.pickFile.addEventListener("click", () => els.priceFile.click());
  els.priceFile.addEventListener("change", async () => {
    const file = els.priceFile.files[0];
    if (!file) return;
    try {
      await loadPriceFile(file);
    } catch (error) {
      els.statusText.textContent = "No se pudo cargar la lista de precios.";
      toast(error.message);
    } finally {
      els.priceFile.value = "";
    }
  });

  els.search.addEventListener("input", filterProducts);
  els.productsBody.addEventListener("click", (event) => {
    const button = event.target.closest(".row-add");
    if (!button) return;
    addProduct(state.filteredProducts[Number(button.dataset.index)]);
  });
  els.quoteList.addEventListener("click", (event) => {
    const button = event.target.closest(".remove");
    if (!button) return;
    state.quoteItems.splice(Number(button.dataset.index), 1);
    renderQuote();
  });
  [els.discount, els.iva, els.extra].forEach((input) => input.addEventListener("input", renderTotals));
  els.exportQuote.addEventListener("click", exportQuote);
  els.clearQuote.addEventListener("click", clearQuote);

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("service-worker.js").catch(() => {});
  }

  renderQuote();
}

setup();
