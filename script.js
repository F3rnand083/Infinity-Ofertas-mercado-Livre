// Lê o Excel .xlsx e devolve um array de produtos
async function loadProductsFromXLSX() {
  const res = await fetch("products.xlsx", { cache: "no-store" });
  if (!res.ok) throw new Error("Falha ao carregar products.xlsx");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const firstSheetName = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  return rows.map(r => ({
    name: (r.name ?? r.Name ?? "").toString().trim(),
    image: (r.image ?? r.Image ?? "").toString().trim(),
    buy_url: (r.buy_url ?? r.Buy_URL ?? r.link ?? "").toString().trim(),
    category: (r.category ?? r.Category ?? "").toString().trim(),
  })).filter(p => p.name && (p.image || p.buy_url));
}

function createCard(product) {
  const card = document.createElement("article");
  card.className = "card";
  const img = document.createElement("img");
  img.alt = product.name || "Produto";
  img.loading = "lazy";
  img.src = product.image || "";
  img.onerror = () => {
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(
      `<svg xmlns='http://www.w3.org/2000/svg' width='600' height='400'>
         <rect width='100%' height='100%' fill='#0b1020'/>
         <text x='50%' y='50%' fill='#94a3b8' font-size='22' text-anchor='middle' dominant-baseline='middle'>Imagem indisponível</text>
       </svg>`
    );
  };
  const body = document.createElement("div");
  body.className = "card-body";
  const title = document.createElement("h3");
  title.className = "card-title";
  title.textContent = product.name || "Sem nome";
  const meta = document.createElement("div");
  meta.className = "card-meta";
  if (product.category) {
    const cat = document.createElement("span");
    cat.className = "badge";
    cat.textContent = product.category;
    meta.appendChild(cat);
  }
  const actions = document.createElement("div");
  actions.className = "card-actions";
  const buy = document.createElement("a");
  buy.className = "btn-buy";
  buy.href = product.buy_url || "#";
  buy.target = "_blank";
  buy.rel = "noopener";
  buy.textContent = "Visite a melhor oferta aqui";
  actions.appendChild(buy);
  body.append(title, meta, actions);
  card.append(img, body);
  return card;
}

function render(products) {
  const grid = document.getElementById("productGrid");
  grid.innerHTML = "";
  products.forEach(p => grid.appendChild(createCard(p)));
}

function sortProducts(products, mode) {
  const cp = [...products];
  switch (mode) {
    case "name-asc":  cp.sort((a,b) => (a.name||"").localeCompare(b.name||"")); break;
    case "name-desc": cp.sort((a,b) => (b.name||"").localeCompare(a.name||"")); break;
  }
  return cp;
}

function filterProducts(products, term, category) {
  const t = (term || "").trim().toLowerCase();
  const c = (category || "").trim().toLowerCase();
  return products.filter(p => {
    const matchesTerm = !t || (p.name || "").toLowerCase().includes(t);
    const matchesCat = !c || (p.category || "").toLowerCase() === c;
    return matchesTerm && matchesCat;
  });
}

function buildCategoryOptions(products) {
  const select = document.getElementById("categorySelect");
  const unique = Array.from(new Set(products.map(p => p.category).filter(Boolean)))
    .sort((a,b) => a.localeCompare(b));
  select.querySelectorAll("option:not(:first-child)").forEach(o => o.remove());
  unique.forEach(cat => {
    const opt = document.createElement("option");
    opt.value = cat;
    opt.textContent = cat;
    select.appendChild(opt);
  });
}

(async function init() {
  document.getElementById("year").textContent = new Date().getFullYear();
  let products = [];
  try {
    products = await loadProductsFromXLSX();
  } catch (e) {
    console.error(e);
    alert("Não foi possível carregar a planilha de produtos (XLSX).");
  }
  buildCategoryOptions(products);
  const searchInput = document.getElementById("searchInput");
  const sortSelect = document.getElementById("sortSelect");
  const categorySelect = document.getElementById("categorySelect");
  function update() {
    const filtered = filterProducts(products, searchInput.value, categorySelect.value);
    const sorted = sortProducts(filtered, sortSelect.value);
    render(sorted);
  }
  searchInput.addEventListener("input", update);
  sortSelect.addEventListener("change", update);
  categorySelect.addEventListener("change", update);
  update();
})();
