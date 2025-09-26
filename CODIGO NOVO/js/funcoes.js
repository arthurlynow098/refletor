const baseSheetStructure = Array.from(
  {
    length: 200,
  },
  (_, i) => ({
    row: i + 1,
    ms: "0",
    values: Array(24).fill(""),
  })
);

// === FUNÇÕES DE INICIALIZAÇÃO E LAYOUT ===
async function initializeApp() {
  assignDomElements();
  injectDynamicStyles();
  initializeDataStructure();
  createSheetSelectorTabs();
  createTableHeader(); // 1. PRIMEIRO: Espera a lista de todos os layouts ser carregada e preparada.
  allAvailableLayouts = await populateLayoutPanel(); // 2. SEGUNDO: Agora que a lista existe, verifica se há um template para carregar.
  const templateDataJSON = localStorage.getItem("simulation_template");
  let layoutFoiCarregadoPeloTemplate = false;

  if (templateDataJSON) {
    try {
      const loadedData = JSON.parse(templateDataJSON); // A função loadDataIntoUI é chamada aqui, com a certeza de que allAvailableLayouts está preenchida.
      layoutFoiCarregadoPeloTemplate = loadDataIntoUI(loadedData);
      localStorage.removeItem("simulation_template");
    } catch (error) {
      console.error("Falha ao processar template do localStorage:", error);
      alert("Ocorreu um erro ao tentar carregar o template.");
    }
  } // 3. TERCEIRO: Se nenhum template foi carregado, carrega o layout padrão.
  // ALTERAÇÃO: Não carrega mais um layout padrão ao iniciar.
  // A aplicação começará com o placeholder visível.
  renderTable(allMacroData[activeSheet]);
  // 4. QUARTO: Continua com o resto da inicialização.
  updateActiveSheetUI();
  populateAdvancedSettingsTables(); // Popula as tabelas de definições
  initializeEventListeners();
  initializeInteractionHandlers();
  renderAllActiveSettings();
  updateConsumptionDisplay();
}

function assignDomElements() {
  configModal = document.getElementById("config-modal");
  openModalBtn = document.getElementById("edit-config-btn");
  closeModalBtn = document.getElementById("close-modal-btn");
  saveAndCloseModalBtn = document.getElementById("save-and-close-modal-btn");
  executeMacroBtn = document.getElementById("execute-macro-btn");
  blocksContainer = document.getElementById("blocks-container");
  macroTableBody = document.getElementById("macro-body");
  macroHeader = document.getElementById("macro-header");
  sheetSelectorContainer = document.getElementById("sheet-selector-container");
  gravarBtn = document.getElementById("gravar-btn");
  restaurarBtn = document.getElementById("restaurar-btn");
  limpaBtn = document.getElementById("limpa-btn");
  colorContextMenu = document.getElementById("color-context-menu");
  layoutPanelModal = document.getElementById("layout-panel-modal");
  openLayoutPanelBtn = document.getElementById("open-layout-panel-btn");
  closeLayoutPanelBtn = document.getElementById("close-layout-panel-btn");
  layoutPanelBody = document.getElementById("layout-panel-body");
  currentLayoutNameSpan = document.getElementById("current-layout-name");
  advancedSettingsModal = document.getElementById("advanced-settings-modal");
  openAdvancedSettingsBtn = document.getElementById("advanced-settings-btn");
  closeAdvancedSettingsBtn = document.getElementById(
    "close-advanced-modal-btn"
  );
  saveAdvancedSettingsBtn = document.getElementById(
    "save-advanced-settings-btn"
  );
  tabButtons = document.querySelectorAll(".modal-tab-btn");
  tabContents = document.querySelectorAll(".modal-tab-content");
  jumpTabelaCheckboxes = document.querySelectorAll(
    '#tab-jumptabela input[type="checkbox"]'
  );
  specialFunctionCells = document.querySelectorAll(
    "#tab-special-functions td[data-block]"
  );
  jumpTabelaStatusContainer = document.getElementById("jumptabela-active-list");
  specialFunctionsStatusContainer = document.getElementById(
    "special-functions-active-list"
  );
  simulationArea = document.getElementById("simulation-area");
  editLayoutBtn = document.getElementById("edit-layout-btn");
  saveLayoutBtn = document.getElementById("save-layout-btn");
}

/**
 * Injeta estilos CSS dinamicamente no <head> do documento.
 * Isso evita a necessidade de modificar o arquivo HTML para adicionar novas cores ou estilos.
 */
function injectDynamicStyles() {
  const style = document.createElement("style");
  style.type = "text/css";
  style.innerHTML = `
    .cell-green {
      background-color: #2ecc71 !important; /* Verde Esmeralda */
      color: white;
      font-weight: bold;
    }
    .block-placeholder {
      background-color: #6c6c6c !important; /* Cor cinza para o fundo */
      border: 1px solid #bdbdbd;
      z-index: 1; /* Garante que fique atrás dos blocos de canais */
      pointer-events: none; /* Impede qualquer interação com o mouse */
    }
    .block-placeholder .block-part-single,
    .block-placeholder .block-part-container {
      background-color: transparent !important;
      color: transparent !important; /* Esconde o texto */
    }
    .block-container {
      z-index: 2; /* Garante que os blocos de canal fiquem na frente */
    }
  `;
  document.head.appendChild(style);
}

async function populateLayoutPanel() {
  layoutPanelBody.innerHTML = "Carregando layouts...";

  // CORREÇÃO: A variável correta é 'AVAILABLE_LAYOUTS' do arquivo mapaBarras.js
  let baseLayouts = JSON.parse(JSON.stringify(layouts));

  try {
    const response = await fetch("carregar_layouts.php");
    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const savedLayouts = await response.json();

    const savedLayoutsMap = new Map();
    savedLayouts.forEach((saved) => {
      savedLayoutsMap.set(saved.layout_id, saved);
    });

    baseLayouts.forEach((layout) => {
      if (savedLayoutsMap.has(layout.id)) {
        const savedData = savedLayoutsMap.get(layout.id);
        layout.blocks = savedData.dados_layout;
      }
    });
  } catch (error) {
    console.error("Não foi possível carregar personalizações:", error);
  }

  renderLayoutGrid(baseLayouts);
  return baseLayouts;
}

function renderLayoutGrid(layoutsToRender) {
  const layoutGroups = layoutsToRender.reduce((groups, layout) => {
    const category = layout.category || "Outros Modelos";
    if (!groups[category]) groups[category] = [];
    groups[category].push(layout);
    return groups;
  }, {});

  const categoryOrder = [
    "Padrão",
    "40 Polegadas",
    "47 Polegadas",
    "54 Polegadas",
    "61 Polegadas",
    "Venus",
    "Outros Modelos",
  ];
  const gridContainer = document.createElement("div");
  gridContainer.className = "layout-panel-grid";

  categoryOrder.forEach((categoryName) => {
    if (layoutGroups[categoryName]) {
      const title = document.createElement("h3");
      title.className = "layout-category-title";
      title.textContent = categoryName;
      gridContainer.appendChild(title);

      layoutGroups[categoryName].forEach((layout) => {
        const item = document.createElement("div");
        item.className = "layout-preview-item";
        item.dataset.layoutId = layout.id;
        item.innerHTML = `<img src="${layout.image}" alt="${layout.name}"><p>${layout.name}</p>`;
        item.addEventListener("click", () => {
          switchLayout(layout.id, layoutsToRender);
          layoutPanelModal.style.display = "none";
        });
        gridContainer.appendChild(item);
      });
    }
  });
  layoutPanelBody.innerHTML = "";
  layoutPanelBody.appendChild(gridContainer);
}

function switchLayout(layoutId, allLayoutsList) {
  if (!allLayoutsList) {
    console.error("A lista de layouts não foi fornecida para switchLayout.");
    populateLayoutPanel().then((layouts) => switchLayout(layoutId, layouts));
    return;
  }

  const selectedLayout = allLayoutsList.find((l) => l.id === layoutId);
  if (!selectedLayout) {
    console.error("Layout não encontrado:", layoutId);
    return;
  }

  // Oculta o placeholder quando um layout é carregado
  const placeholder = document.getElementById("simulation-placeholder");
  if (placeholder) {
    placeholder.style.display = "none";
  }
  currentLayout = selectedLayout;
  activeBlockLayout = JSON.parse(JSON.stringify(selectedLayout.blocks));
  currentLayoutNameSpan.textContent = currentLayout.name;
  document.getElementById("siren-model-image").src = currentLayout.image;
  initializeBlocks(activeBlockLayout);
  stopSimulation();
}

/**
 * Gera dinamicamente o conteúdo das tabelas de Funções Especiais.
 */
function populateAdvancedSettingsTables() {
  const functionGroups = [
    {
      instruction: "Apagar Blocos",
      func: "takedown",
      tbodyId: "takedown-off-tbody",
    },
    {
      instruction: "Acender Blocos",
      func: "takedown_on",
      tbodyId: "takedown-on-tbody",
    },
    {
      instruction: "Apagar Blocos",
      func: "alley_right",
      tbodyId: "alley-right-off-tbody",
    },
    {
      instruction: "Acender Blocos",
      func: "alley_right_on",
      tbodyId: "alley-right-on-tbody",
    },
    {
      instruction: "Apagar Blocos",
      func: "alley_left",
      tbodyId: "alley-left-off-tbody",
    },
    {
      instruction: "Acender Blocos",
      func: "alley_left_on",
      tbodyId: "alley-left-on-tbody",
    },
    {
      instruction: "Apagar Blocos",
      func: "backlight_off",
      tbodyId: "backlight-off-tbody",
    },
    {
      instruction: "Acender Blocos",
      func: "backlight_on",
      tbodyId: "backlight-on-tbody",
    },
    {
      instruction: "Apagar Blocos",
      func: "cut_front_off",
      tbodyId: "cut-front-off-tbody",
    },
    {
      instruction: "Acender Blocos",
      func: "cut_front_on",
      tbodyId: "cut-front-on-tbody",
    },
    {
      instruction: "Apagar Blocos",
      func: "cut_rear_off",
      tbodyId: "cut-rear-off-tbody",
    },
    {
      instruction: "Acender Blocos",
      func: "cut_rear_on",
      tbodyId: "cut-rear-on-tbody",
    },
    {
      instruction: "Apagar Blocos",
      func: "cut_dt_off",
      tbodyId: "cut-dt-off-tbody",
    },
    {
      instruction: "Acender Blocos",
      func: "cut_dt_on",
      tbodyId: "cut-dt-on-tbody",
    },
    {
      instruction: "Apagar Blocos",
      func: "cut_horn_light",
      tbodyId: "cut-horn-light-tbody",
    },
  ];

  document.querySelectorAll(".function-table").forEach((table) => {
    const thead = table.querySelector("thead tr");
    if (!thead) return;
    // Limpa cabeçalho antigo e gera novo
    thead.innerHTML = "<th>Instrução</th>";
    for (let i = 1; i <= 24; i++) {
      thead.innerHTML += `<th>${i}</th>`;
    }
  });

  functionGroups.forEach(({ instruction, func, tbodyId }) => {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    let cells = "";
    for (let i = 1; i <= 24; i++) {
      cells += `<td contenteditable="true" data-function="${func}" data-block="${i}"></td>`;
    }
    tbody.innerHTML = `<tr><td>${instruction}</td>${cells}</tr>`;
  });

  // Popula os selects
  const numFpSelect = document.getElementById("config-num-fp-principais");
  const qtdFpDsSelect = document.getElementById("config-qtd-fp-ds");
  for (let i = 1; i <= 15; i++) {
    numFpSelect.innerHTML += `<option value="${i}" ${
      i === 4 ? "selected" : ""
    }>${i}</option>`;
    qtdFpDsSelect.innerHTML += `<option value="${i}" ${
      i === 8 ? "selected" : ""
    }>${i}</option>`;
  }
}

function initializeDataStructure() {
  sheetNames.forEach((name) => {
    allMacroData[name] = JSON.parse(JSON.stringify(baseSheetStructure));
    undoHistory[name] = [];
    redoHistory[name] = [];
  });
  columnColors = {};
}

function initializeBlocks(blocksArray) {
  blocksContainer.innerHTML = "";
  if (!blocksArray) return;
  blocksArray.forEach((blockData, index) => {
    const blockElement = document.createElement("div");
    blockElement.className = "block-container";
    blockElement.dataset.index = index;
    blockElement.style.top = blockData.top;
    blockElement.style.left = blockData.left;
    blockElement.style.width = `${blockData.width}px`;
    blockElement.style.height = `${blockData.height}px`;
    blockElement.style.transform = `rotate(${blockData.rotate}deg)`;

    if (blockData.type === "placeholder") {
      blockElement.classList.add("block-placeholder");
      blockData.text = ""; // Garante que o placeholder não tenha texto.
    }

    // --- CORREÇÃO PRINCIPAL: Ouvir o duplo clique no CONTAINER, não na parte interna ---
    blockElement.addEventListener("dblclick", (e) => {
      // 1. Impede a propagação e o comportamento padrão IMEDIATAMENTE
      e.preventDefault();
      e.stopPropagation();

      // 2. Só funciona se o modo de edição estiver ativo
      if (!isEditMode || blockData.type === "placeholder") return;

      // 3. Encontra a parte interna do bloco (o elemento que contém o texto)
      const singlePart = blockElement.querySelector(".block-part-single");
      if (!singlePart) {
        console.warn("Tentativa de editar um bloco que não é de parte única.");
        return;
      }

      // 4. Guarda o valor original e torna o campo editável
      const originalValue = singlePart.textContent;
      singlePart.contentEditable = true;
      singlePart.focus();
      document.execCommand("selectAll", false, null);

      // 5. Define a função que será chamada ao sair do campo (blur)
      const handleBlur = () => {
        const blockContainer = singlePart.closest(".block-container");
        if (!blockContainer) return;
        const dataIndex = parseInt(blockContainer.dataset.index, 10);
        const newText = singlePart.textContent.trim();
        const newNumber = parseInt(newText, 10);

        // 6. VALIDAÇÃO: Verifica se é um número válido entre 1 e 24
        if (!isNaN(newNumber) && newNumber >= 1 && newNumber <= 24) {
          activeBlockLayout[dataIndex].text = String(newNumber);
          singlePart.dataset.blockNum = String(newNumber);
          singlePart.textContent = String(newNumber);
        } else {
          // Se for inválido, mostra um alerta e restaura o valor original
          alert("Entrada inválida. Por favor, insira um número entre 1 e 24.");
          singlePart.textContent = originalValue;
        }

        // 7. Finaliza a edição e remove os listeners
        singlePart.contentEditable = false;
        singlePart.removeEventListener("blur", handleBlur);
        singlePart.removeEventListener("keydown", handleKeyDown);
      };

      // 8. Define a função para lidar com as teclas Enter e Escape
      const handleKeyDown = (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          singlePart.blur(); // Aciona o handleBlur para salvar
        } else if (e.key === "Escape") {
          singlePart.textContent = originalValue; // Restaura o valor
          singlePart.blur(); // Aciona o handleBlur para cancelar e limpar
        } else if (e.key === "Backspace" || e.key === "Delete") {
          // Permite a exclusão de caracteres
          // Nenhuma ação adicional é necessária aqui, o navegador trata nativamente.
          // A validação acontecerá apenas no blur.
        }
      };

      // 9. Ativa os listeners de blur e keydown
      singlePart.addEventListener("blur", handleBlur);
      singlePart.addEventListener("keydown", handleKeyDown);
    });
    // --- FIM DA CORREÇÃO PRINCIPAL ---

    if (Array.isArray(blockData.text)) {
      // Bloco dividido (Bicolor)
      const container = document.createElement("div");
      container.className = "block-part-container";
      const topPart = document.createElement("div");
      topPart.className = "block-part-top layout-color-off";
      topPart.textContent = blockData.text[0];
      topPart.dataset.blockNum = blockData.text[0];
      const bottomPart = document.createElement("div");
      bottomPart.className = "block-part-bottom layout-color-off";
      bottomPart.textContent = blockData.text[1];
      bottomPart.dataset.blockNum = blockData.text[1];
      container.appendChild(topPart);
      container.appendChild(bottomPart);
      blockElement.appendChild(container);
    } else {
      // Bloco único
      const singlePart = document.createElement("div");
      singlePart.className = "block-part-single layout-color-off";
      singlePart.textContent = blockData.text;
      singlePart.dataset.blockNum = blockData.text;
      blockElement.appendChild(singlePart);
    }

    blockElement.innerHTML += `<div class="interaction-handle resize-handle"></div><div class="interaction-handle rotate-handle"></div>`;
    blocksContainer.appendChild(blockElement);
  });
}

function loadDataIntoUI(loadedData) {
  console.log("--- INICIANDO O CARREGAMENTO DO TEMPLATE ---");
  console.log("1. Objeto completo recebido do localStorage:", loadedData);

  let layoutSwitched = false;
  if (loadedData && loadedData.layout_id) {
    const idDoLayoutSalvo = loadedData.layout_id;
    console.log(
      "2. ID do layout encontrado no template:",
      `'${idDoLayoutSalvo}'`
    );
    console.log(
      "3. Procurando este ID na lista de layouts disponíveis:",
      allAvailableLayouts
    );

    const foundLayout = allAvailableLayouts.find(
      (l) => l.id === idDoLayoutSalvo
    );
    console.log("4. Resultado da busca (Layout Encontrado):", foundLayout);

    if (foundLayout) {
      console.log("5. SUCESSO! Trocando para o layout:", foundLayout.name);
      switchLayout(idDoLayoutSalvo, allAvailableLayouts);
      layoutSwitched = true;
    } else {
      console.error(
        "5. FALHA! O ID do layout salvo não corresponde a nenhum layout conhecido."
      );
    }
  } else {
    console.error("O objeto carregado não possui a chave 'layout_id'.");
  } // O resto da função continua para carregar os dados da macro, cores, etc.

  if (loadedData.block_layout && !layoutSwitched) {
    activeBlockLayout = loadedData.block_layout;
    if (currentLayout) {
      currentLayout.blocks = JSON.parse(JSON.stringify(activeBlockLayout));
    }
    initializeBlocks(activeBlockLayout);
  }

  allMacroData = loadedData.data || {};
  sheetNames.forEach((name) => {
    if (!allMacroData[name]) {
      allMacroData[name] = JSON.parse(JSON.stringify(baseSheetStructure));
    }
  });
  savedSpecialFunctionConfigs = loadedData.special_functions || {};
  savedJumpTabelaSettings = loadedData.jumptabela_settings || {};
  columnColors = loadedData.column_colors || {};
  specialFunctionCells.forEach((cell) => {
    const func = cell.dataset.function;
    const block = parseInt(cell.dataset.block, 10);
    cell.textContent = savedSpecialFunctionConfigs[func]?.includes(block)
      ? "1"
      : "";
  });
  jumpTabelaCheckboxes.forEach((checkbox) => {
    const key = `jt${checkbox.dataset.jumptabela}_bit${checkbox.dataset.bit}`;
    checkbox.checked = !!savedJumpTabelaSettings[key];
  });
  activeSheet = "FP1";
  renderTable(allMacroData[activeSheet]);
  updateActiveSheetUI();
  updateAllVisuals();
  renderAllActiveSettings();
  updateConsumptionDisplay();
  alert("Template carregado com sucesso!");
  return layoutSwitched;
}

function pushStateToUndoHistory() {
  const currentData = getCurrentTableData();
  const lastState =
    undoHistory[activeSheet][undoHistory[activeSheet].length - 1];
  if (lastState && JSON.stringify(lastState) === JSON.stringify(currentData)) {
    return;
  }
  undoHistory[activeSheet].push(JSON.parse(JSON.stringify(currentData)));
  redoHistory[activeSheet] = [];
}

const debouncedPushState = () => {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => {
    pushStateToUndoHistory();
  }, 500);
};

function createSheetSelectorTabs() {
  sheetSelectorContainer.innerHTML = "";
  sheetNames.forEach((name) => {
    const button = document.createElement("button");
    button.className = "sheet-selector-btn";
    button.dataset.sheet = name;
    button.textContent = name;
    sheetSelectorContainer.appendChild(button);
  });
}

function createTableHeader() {
  macroHeader.innerHTML = "";
  const tr = document.createElement("tr");
  tr.innerHTML += `<th class="column-header">ms/25</th><th class="column-header">ms</th><th class="column-header">#</th>`;
  for (let i = 1; i <= 24; i++) {
    tr.innerHTML += `<th class="column-header" data-col-index="${
      i - 1
    }" title="Clique com o botão direito para definir uma cor fixa para este canal.">${i}</th>`;
  }
  macroHeader.appendChild(tr);
}

function renderTable(data) {
  macroTableBody.innerHTML = "";
  const tableData = data || JSON.parse(JSON.stringify(baseSheetStructure));
  tableData.forEach((rowData) => {
    const tr = document.createElement("tr");
    const ms25Cell = document.createElement("td");
    ms25Cell.className = "ms25-column";
    ms25Cell.contentEditable = true;
    const msValue = parseInt(rowData.ms, 10);
    ms25Cell.textContent =
      isNaN(msValue) || rowData.ms === ""
        ? ""
        : msValue === 0
        ? "0"
        : Math.round(msValue / 25);
    tr.appendChild(ms25Cell);
    const msCell = document.createElement("td");
    msCell.className = "ms-column";
    msCell.textContent = rowData.ms;
    msCell.contentEditable = true;
    tr.appendChild(msCell);
    const counterCell = document.createElement("td");
    counterCell.className = "row-counter-column";
    counterCell.textContent = rowData.row;
    tr.appendChild(counterCell);
    ms25Cell.addEventListener("input", () =>
      syncTimeValues(ms25Cell, msCell, "from_ms25")
    );
    msCell.addEventListener("input", () =>
      syncTimeValues(ms25Cell, msCell, "from_ms")
    );
    const dataValues = rowData.values || [];
    for (let i = 0; i < 24; i++) {
      createEditableCell(tr, dataValues[i] || "", i);
    }
    macroTableBody.appendChild(tr);
  });
  applyAllColumnColors();
}

function syncTimeValues(ms25Cell, msCell, source) {
  if (source === "from_ms25") {
    const val = parseInt(ms25Cell.textContent, 10);
    msCell.textContent = isNaN(val) ? "" : val * 25;
  } else {
    const val = parseInt(msCell.textContent, 10);
    ms25Cell.textContent = isNaN(val) ? "" : Math.round(val / 25);
  }
}

function createEditableCell(parent, value, colIndex) {
  const td = document.createElement("td");
  td.className = "editable-cell";
  td.contentEditable = true;
  td.dataset.colIndex = colIndex;
  td.textContent = value;
  updateCellColor(td);
  td.addEventListener("blur", pushStateToUndoHistory);
  td.addEventListener("input", function () {
    // Qualquer número digitado vira "1". Se apagar, fica vazio.
    const originalText = this.textContent;
    const hasNumber = /[1-9]/.test(originalText);
    const newText = hasNumber ? "1" : "";
    if (originalText !== newText) {
      this.textContent = newText;
    }
    updateCellColor(this);
    updateConsumptionDisplay();
    debouncedPushState();
  });
  td.addEventListener("focus", function () {
    highlightCurrentRow(this.parentElement.rowIndex - 1);
  });
  parent.appendChild(td);
}

function updateCellColor(cell) {
  // Esta função agora apenas garante que a cor da coluna seja aplicada ou removida.
  const ALL_COLOR_CLASSES = [
    "cell-red",
    "cell-blue",
    "cell-green",
    "cell-yellow",
    "cell-white",
    "persistent-color-red",
    "persistent-color-green",
    "persistent-color-blue",
    "persistent-color-white",
    "persistent-color-yellow",
  ];
  const value = cell.textContent.trim();
  const colIndex = cell.dataset.colIndex;
  const persistentColor = columnColors[colIndex];

  // 1. Limpa todas as classes de cor para evitar conflitos.
  cell.classList.remove(...ALL_COLOR_CLASSES);

  if (value === "1") {
    // 2. Se a célula tem valor "1", aplica a cor sólida (cell-xxxx).
    if (persistentColor) {
      // Mapeia a cor da coluna para a classe de cor sólida.
      if (persistentColor === "red") cell.classList.add("cell-red");
      else if (persistentColor === "blue") cell.classList.add("cell-blue");
      else if (persistentColor === "white") cell.classList.add("cell-white");
      else if (persistentColor === "yellow") cell.classList.add("cell-yellow");
      else if (persistentColor === "green") cell.classList.add("cell-green");
    } else {
      // Se a coluna não tem cor, usa vermelho como padrão para o valor "1".
      cell.classList.add("cell-red");
    }
  } else {
    // 3. Se a célula está vazia, aplica a cor pastel da coluna (persistent-color-xxxx).
    if (persistentColor) {
      cell.classList.add(`persistent-color-${persistentColor}`);
    }
  }
}

// ========================================================================================//
// =================== FUNÇÃO PARA DETECTAR LAYOUT BASEADO NOS DADOS ================== //
// ========================================================================================//
function detectLayoutFromData(extractedData, fallbackData = null) {
  // Usa dados do Python se disponível, senão usa dados do JavaScript
  const data = extractedData || fallbackData;

  if (!data) return null;

  let modules = [];
  let productInfo = {};
  let totalModules = 0;

  // Se temos dados do Python (formato completo)
  if (data.mapeamento_cores_modulos || data.pythonData) {
    const pythonData = data.pythonData || data;

    // Conta módulos ativos do mapeamento
    if (pythonData.mapeamento_cores_modulos) {
      modules = pythonData.mapeamento_cores_modulos
        .filter((mod) => mod.mapeamento && mod.mapeamento.length > 0)
        .map((mod) => ({
          module: mod.modulo,
          color: mod.mapeamento[0]?.cor || "vermelho",
        }));
    }

    totalModules = pythonData.quantidade_modulos_ativos || modules.length;

    productInfo = {
      produto: pythonData.produto || "",
      tecnologia: pythonData.tecnologia || pythonData.modelo || "",
      tamanho_polegadas: pythonData.tamanho_polegadas || "",
    };
  }
  // Se temos dados do JavaScript (formato simples)
  else if (data.layout || data.modules) {
    modules = data.layout || data.modules || [];
    totalModules = modules.length;
    productInfo = data.product_info || {
      produto: "",
    };
  }

  if (totalModules === 0) return null;

  const produto = (productInfo.produto || "").toLowerCase();
  const tecnologia = (productInfo.tecnologia || "").toLowerCase();
  const tamanho = productInfo.tamanho_polegadas || "";

  console.log(
    `Detectando layout: ${totalModules} módulos ativos, produto: ${produto}, tecnologia: ${tecnologia}, tamanho: ${tamanho}P`
  );

  // Detecta tipo de produto baseado nos dados extraídos
  let productType = "";
  if (produto.includes("primus")) {
    if (tamanho) {
      productType = `${tamanho}p`;
    }
  } else if (produto.includes("venus")) {
    productType = "venus";
  }

  // CORRIGIDO: ECO e STAR são Colimadores, o resto é Refletor.
  const isCollimator =
    tecnologia.includes("eco") || tecnologia.includes("star");
  const layoutSuffix = isCollimator ? "Col" : "Ref";

  // Gera o ID exato do layout esperado
  let exactLayoutId = "";
  if (productType === "venus") {
    exactLayoutId = `venus${totalModules}m`;
  } else {
    exactLayoutId = `primus${productType}${totalModules}m${layoutSuffix}`;
  }

  // Verifica se o layout exato existe na lista de layouts disponíveis
  const layoutExists = allAvailableLayouts.some((l) => l.id === exactLayoutId);

  if (layoutExists) {
    console.log(`✓ Layout exato detectado: ${exactLayoutId}`);
    return exactLayoutId;
  }

  // Se não existir, retorna null para indicar que a combinação é inválida
  console.error(
    `❌ Layout inválido! A combinação "${exactLayoutId}" não existe.`
  );
  return null;
}

// ========================================================================================//
// =================== FUNÇÃO DE EXTRAÇÃO DO EXCEL (PYTHON BACKEND) ================= //
// ========================================================================================//
async function processExcelLayout(file) {
  try {
    console.log("Enviando arquivo para processamento Python...");

    // Cria FormData para enviar o arquivo
    const formData = new FormData();
    formData.append("excel_file", file);

    // Envia para o backend PHP que chama o Python
    const response = await fetch("process_excel.php", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      throw new Error(`Erro HTTP: ${response.status} - ${response.statusText}`);
    }

    const contentType = response.headers.get("content-type");
    if (!contentType || !contentType.includes("application/json")) {
      const textResponse = await response.text();
      console.error("Resposta não é JSON:", textResponse);
      throw new Error("Servidor retornou resposta inválida (não JSON)");
    }

    const result = await response.json();
    console.log("✓ Processamento Python bem-sucedido!", result);

    if (result.error) {
      throw new Error(`Erro do Python: ${result.error}`);
    }

    // Se o PHP retorna um wrapper com 'data', extrai os dados
    const extractedData = result.data || result;
    return extractedData;
  } catch (error) {
    console.error("Erro ao processar Excel com Python:", error);

    // Se falhar, usa implementação JavaScript como fallback
    console.log("🔄 Tentando com implementação JavaScript como fallback...");
    try {
      const jsResult = await processExcelLayoutJS(file);
      console.log("✓ Processamento JavaScript bem-sucedido!", jsResult);
      return jsResult;
    } catch (jsError) {
      console.error("Erro também no processamento JavaScript:", jsError);
      throw new Error(
        `Ambos os processamentos falharam. Python: ${error.message}, JavaScript: ${jsError.message}`
      );
    }
  }
}

// Implementação JavaScript pura como fallback (formato Python)
async function processExcelLayoutJS(file) {
  const COLOR_MAP = {
    VM: "vermelho",
    AZ: "azul",
    BR: "branco",
    AB: "ambar",
  };
  const VALID_MODULE_NUMBERS = Array.from(
    {
      length: 28,
    },
    (_, i) => i + 1
  );

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Erro ao ler o arquivo."));

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, {
          type: "array",
        });
        const sheetNames = workbook.SheetNames;

        // --- LÓGICA MELHORADA PARA SELECIONAR A PLANILHA E O TAMANHO ---
        let targetSheetName = null;
        let tamanhoDetectadoPelaPlanilha = null;
        const tamanhosPossiveis = ["40", "47", "54", "61"];

        // 1. Procura por uma planilha que corresponda exatamente a um dos tamanhos
        for (const tamanho of tamanhosPossiveis) {
          const nomePlanilhaProcurado = `${tamanho} POL`;
          const foundSheet = sheetNames.find(
            (name) => name.toUpperCase() === nomePlanilhaProcurado
          );
          if (foundSheet) {
            targetSheetName = foundSheet;
            tamanhoDetectadoPelaPlanilha = tamanho;
            console.log(
              `Planilha encontrada para o tamanho: ${tamanhoDetectadoPelaPlanilha}`
            );
            break;
          }
        }

        // 2. Se não encontrar, usa a lógica de fallback (primeira com "POL" ou a primeira de todas)
        if (!targetSheetName) {
          targetSheetName =
            sheetNames.find((name) => name.toUpperCase().includes("POL")) ||
            sheetNames[0];
          console.log(
            `Nenhuma planilha de tamanho específico encontrada. Usando fallback: ${targetSheetName}`
          );
        }
        // --- FIM DA LÓGICA MELHORADA ---

        const worksheet = workbook.Sheets[targetSheetName];
        if (!worksheet)
          return reject(
            new Error(`A planilha '${targetSheetName}' não foi encontrada.`)
          );

        // Busca informações do produto (como o Python faz)
        let productInfo = {
          codigo_firmware: "",
          produto: "",
          tecnologia: "", // Renomeado de 'modelo' para 'tecnologia'
          cores_leds: "()",
          funcoes_extras: "",
          cores_tampas: "",
          tamanho_polegadas: "",
          quantidade_modulos: "0",
          numero_leds: "0",
          cores_leds_lista: [],
        };

        // Procura pela célula com "SIN VISUAL" para extrair informações do produto
        console.log(
          `Buscando informações do produto na planilha: ${targetSheetName}`
        );

        // Obtém o range real da planilha
        const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:Z100");
        console.log(
          `Range real da planilha: ${worksheet["!ref"] || "A1:Z100"} (${
            range.e.r + 1
          } linhas, ${range.e.c + 1} colunas)`
        );
        console.log(
          "Primeiras células disponíveis na planilha:",
          Object.keys(worksheet)
            .filter((key) => !key.startsWith("!"))
            .slice(0, 20)
        );

        let infoEncontrada = false;
        let celulasProcessadas = 0;
        let celulasComConteudo = 0;

        // Varre toda a planilha usando o range real
        for (let row = range.s.r + 1; row <= range.e.r + 1; row++) {
          for (let col = range.s.c + 1; col <= range.e.c + 1; col++) {
            celulasProcessadas++;
            const cellAddress = XLSX.utils.encode_cell({
              r: row - 1,
              c: col - 1,
            });
            const cell = worksheet[cellAddress];

            if (cell && cell.v) {
              // CORREÇÃO: Usar cell.v em vez de cell.value
              celulasComConteudo++;
              const cellValue = String(cell.v);

              // Log células nas primeiras linhas para debug
              if (row <= 15 && col <= 15 && cellValue.trim()) {
                // console.log(`Célula ${cellAddress}: "${cellValue}"`); // Comentado para não poluir o console
              }

              // Log células que contenham texto relevante
              if (
                cellValue.includes("PRIMUS") ||
                cellValue.includes("ARES") ||
                cellValue.includes("VENUS") ||
                cellValue.includes("SIN")
              ) {
                console.log(
                  `🔍 Célula interessante ${cellAddress}: "${cellValue}"`
                );
              }

              if (cellValue.includes("SIN VISUAL")) {
                const info = cellValue;
                console.log(
                  "✓ Informação SIN VISUAL encontrada em",
                  cellAddress,
                  ":",
                  info
                );
                infoEncontrada = true;

                // ✅ REGEX CORRIGIDA para o formato: "51.PRM.PM47.00075.00 - SIN VISUAL PRIMUS PRIME 47P 20M 114L (VM+AZ+BR) 6BI TP FM/FM/FM"
                const primusMatch = info.match(
                  /SIN VISUAL\s+(PRIMUS|ARES|VENUS)\s+(ECO|STAR|PRIME|SALIENT|ULTRA)\s+(\d+)P\s+(\d+)M\s+(\d+)L\s+\(([^)]+)\)/i
                );

                if (primusMatch) {
                  console.log("✓ Regex de produto funcionou:", primusMatch);
                  productInfo.produto = primusMatch[1]; // PRIMUS
                  productInfo.tecnologia = primusMatch[2]; // PRIME
                  productInfo.tamanho_polegadas = primusMatch[3]; // 47
                  productInfo.quantidade_modulos = primusMatch[4]; // 20
                  productInfo.numero_leds = primusMatch[5]; // 114
                  productInfo.cores_leds = primusMatch[6]; // VM+AZ+BR
                  productInfo.cores_leds_lista = primusMatch[6].split("+"); // ['VM', 'AZ', 'BR']

                  console.log("✓ Informações extraídas:", {
                    produto: productInfo.produto,
                    tecnologia: productInfo.tecnologia,
                    tamanho: productInfo.tamanho_polegadas,
                    modulos: productInfo.quantidade_modulos,
                    leds: productInfo.numero_leds,
                    cores: productInfo.cores_leds_lista,
                  });
                } else {
                  console.log(
                    "⚠ Regex não funcionou, tentando extração mais robusta..."
                  );

                  // ✅ EXTRAÇÃO MAIS ROBUSTA usando múltiplas regex específicas
                  const produtoMatch = info.match(
                    /SIN VISUAL\s+(PRIMUS|ARES|VENUS)/i
                  );
                  if (produtoMatch) {
                    productInfo.produto = produtoMatch[1];
                    console.log("✓ Produto extraído:", productInfo.produto);
                  }

                  const tecnologiaMatch = info.match(
                    /(ECO|STAR|PRIME|SALIENT|ULTRA)/i
                  );
                  if (tecnologiaMatch) {
                    productInfo.tecnologia = tecnologiaMatch[1];
                    console.log(
                      "✓ Tecnologia extraída:",
                      productInfo.tecnologia
                    );
                  }

                  const tamanhoMatch = info.match(/(\d+)P/);
                  if (tamanhoMatch) {
                    productInfo.tamanho_polegadas = tamanhoMatch[1];
                    console.log(
                      "✓ Tamanho extraído:",
                      productInfo.tamanho_polegadas + "P"
                    );
                  }

                  const modulosMatch = info.match(/(\d+)M/);
                  if (modulosMatch) {
                    productInfo.quantidade_modulos = modulosMatch[1];
                    console.log(
                      "✓ Módulos extraídos:",
                      productInfo.quantidade_modulos + "M"
                    );
                  }

                  const ledsMatch = info.match(/(\d+)L/);
                  if (ledsMatch) {
                    productInfo.numero_leds = ledsMatch[1];
                    console.log(
                      "✓ LEDs extraídos:",
                      productInfo.numero_leds + "L"
                    );
                  }

                  const coresMatch = info.match(/\(([^)]+)\)/);
                  if (coresMatch) {
                    productInfo.cores_leds = coresMatch[1];
                    productInfo.cores_leds_lista = coresMatch[1]
                      .replace(/\+/g, ",")
                      .split(",")
                      .map((c) => c.trim());
                    console.log(
                      "✓ Cores extraídas:",
                      productInfo.cores_leds_lista
                    );
                  }
                }
                break;
              }
            }
          }
          if (infoEncontrada) break;
        }

        console.log(`📊 Estatísticas da varredura:`);
        console.log(`   - Células processadas: ${celulasProcessadas}`);
        console.log(`   - Células com conteúdo: ${celulasComConteudo}`);
        console.log(
          `   - SIN VISUAL encontrado: ${infoEncontrada ? "SIM" : "NÃO"}`
        );

        if (!infoEncontrada) {
          console.log("⚠ Nenhuma célula com 'SIN VISUAL' foi encontrada");
          console.log("💡 Tentando busca mais ampla por palavras-chave...");

          // Busca alternativa por palavras-chave em toda a planilha
          Object.keys(worksheet).forEach((cellKey) => {
            if (!cellKey.startsWith("!")) {
              const cell = worksheet[cellKey]; // Objeto da célula
              if (cell && cell.v) {
                // ✅ CORREÇÃO: Usar 'cell.v' em vez de 'cell.value'
                const cellValue = String(cell.v);
                if (
                  cellValue.includes("PRIMUS") ||
                  cellValue.includes("ARES") ||
                  cellValue.includes("VENUS")
                ) {
                  console.log(
                    `🔍 Palavra-chave encontrada em ${cellKey}: "${cellValue}"`
                  );
                }
              }
            }
          });
        }

        // ========================================================================================//
        // =================== MAPEAMENTO DE LAYOUTS PARA EXTRAÇÃO DO EXCEL ======================= //
        // ========================================================================================//

        const MAPA_LAYOUT_EXCEL = {
          // Mapeamento para 40 Polegadas (a ser preenchido)
          40: {
            posicoes_modulos: [
              "O4:P5",
              "Q4:R5",
              "S4:T5",
              "U6:V7",
              "W10:X11",
              "W14:X15",
              "W18:X19",
              "U22:V23",
              "S24:T25",
              "Q24:R25",
              "O24:P25",
              "L24:M25",
              "J24:K25",
              "H24:I25",
              "F22:G23",
              "D18:E19",
              "D14:E15",
              "D10:E11",
              "F6:G7",
              "H4:I5",
              "J4:K5",
              "L4:M5",
            ],
            posicoes_canais: {
              "O4:P5": [2, 0],
              "Q4:R5": [2, 0],
              "S4:T5": [2, 0],
              "U6:V7": [2, 0],
              "W10:X11": [2, 0],
              "W14:X15": [0, -2],
              "W18:X19": [-2, 0],
              "U22:V23": [-2, 0],
              "S24:T25": [-2, 0],
              "Q24:R25": [-2, 0],
              "O24:P25": [-2, 0],
              "L24:M25": [-2, 0],
              "J24:K25": [-2, 0],
              "H24:I25": [-2, 0],
              "F22:G23": [-2, 0],
              "D18:E19": [-2, 0],
              "D14:E15": [0, 2],
              "D10:E11": [2, 0],
              "F6:G7": [2, 0],
              "H4:I5": [2, 0],
              "J4:K5": [2, 0],
              "L4:M5": [2, 0],
            },
          },
          // Mapeamento para 47 Polegadas (já existente)
          47: {
            posicoes_modulos: [
              "O4:P5",
              "Q4:R5",
              "S4:T5",
              "V4:W5",
              "X4:Y5",
              "Z4:AA5",
              "AB6:AC7",
              "AD10:AE11",
              "AD14:AE15",
              "AD18:AE19",
              "AB22:AC23",
              "Z24:AA25",
              "X24:Y25",
              "V24:W25",
              "S24:T25",
              "Q24:R25",
              "O24:P25",
              "L24:M25",
              "J24:K25",
              "H24:I25",
              "F22:G23",
              "D18:E19",
              "D14:E15",
              "D10:E11",
              "F6:G7",
              "H4:I5",
              "J4:K5",
              "L4:M5",
            ],
            posicoes_canais: {
              "O4:P5": [2, 0],
              "Q4:R5": [2, 0],
              "S4:T5": [2, 0],
              "V4:W5": [2, 0],
              "X4:Y5": [2, 0],
              "Z4:AA5": [2, 0],
              "AB6:AC7": [2, 0],
              "AD10:AE11": [2, 0],
              "AD14:AE15": [0, -2],
              "AD18:AE19": [-2, 0],
              "AB22:AC23": [-2, 0],
              "Z24:AA25": [-2, 0],
              "X24:Y25": [-2, 0],
              "V24:W25": [-2, 0],
              "S24:T25": [-2, 0],
              "Q24:R25": [-2, 0],
              "O24:P25": [-2, 0],
              "L24:M25": [-2, 0],
              "J24:K25": [-2, 0],
              "H24:I25": [-2, 0],
              "F22:G23": [-2, 0],
              "D18:E19": [-2, 0],
              "D14:E15": [0, 2],
              "D10:E11": [2, 0],
              "F6:G7": [2, 0],
              "H4:I5": [2, 0],
              "J4:K5": [2, 0],
              "L4:M5": [2, 0],
            },
          },

          54: {
            posicoes_modulos: [
              "V4:W5",      // 1 - superior, para baixo
    "AD4:AE5",    // 2 - superior, para baixo
    "AG4:AH5",    // 3 - superior, para baixo
    "AM4:AN5",    // 4 - superior, para baixo
    "AM8:AN9",    // 5 - lateral direito, para baixo
    "AM12:AN13",  // 6 - lateral direito, para baixo
    "AM16:AN17",  // 7 - lateral direito, para baixo
    "AM20:AN21",  // 8 - lateral direito, para baixo
    "AG16:AH17",  // 9 - inferior, para baixo
    "W16:X17",    // 10 - inferior, para baixo
    "Q16:R17",    // 11 - inferior, para baixo
    "K16:L17",    // 12 - inferior, para baixo
    "H16:I17",    // 13 - inferior, para baixo
    "D12:E13",    // 14 - lateral esquerdo, para cima
    "D8:E9",      // 15 - lateral esquerdo, para cima
    "D16:E17",    // 16 - lateral esquerdo, para cima
    "D20:E21",    // 17 - lateral esquerdo, para cima
    "H23:I24",    // 18 - inferior, para cima
    "K23:L24",    // 19 - inferior, para cima
    "Q23:R24",    // 20 - inferior, para cima
    "V23:W24",    // 21 - inferior, para cima
    "AD23:AE24",  // 22 - inferior, para cima
    "AG23:AH24",  // 23 - inferior, para cima
    "AM23:AN24" ], // TODO: Preencher
            posicoes_canais: {}, // TODO: Preencher
          },
          // Mapeamento para 61 Polegadas (a ser preenchido)
          61: {
            posicoes_modulos: [], // TODO: Preencher
            posicoes_canais: {}, // TODO: Preencher
          },
        };

        // ========================================================================================//
        // =================== FUNÇÃO DE EXTRAÇÃO DO EXCEL (PYTHON BACKEND) ================= //
        // ========================================================================================//

        // Seleciona o mapeamento correto com base no tamanho em polegadas
        // PRIORIDADE 1: Tamanho detectado pelo nome da planilha.
        // PRIORIDADE 2: Tamanho extraído da descrição do produto.
        // PRIORIDADE 3: Padrão "47".
        const tamanhoFinal =
          tamanhoDetectadoPelaPlanilha || productInfo.tamanho_polegadas || "47";
        const layoutConfig = MAPA_LAYOUT_EXCEL[tamanhoFinal];

        if (!layoutConfig || layoutConfig.posicoes_modulos.length === 0) {
          const erroMsg = `Configuração de layout para ${tamanhoFinal} polegadas não encontrada ou está vazia.`;
          console.error(erroMsg);
          // Opcional: pode rejeitar a promessa aqui se for um erro crítico
          // return reject(new Error(erroMsg));
        }

        const posicoes_modulos =
          layoutConfig?.posicoes_modulos ||
          MAPA_LAYOUT_EXCEL["47"].posicoes_modulos;
        const POSICOES_CANAIS =
          layoutConfig?.posicoes_canais ||
          MAPA_LAYOUT_EXCEL["47"].posicoes_canais;

        const mapeamento_cores_modulos = [];
        console.log("Iniciando processamento de módulos...");

        // Processa cada módulo
        posicoes_modulos.forEach((range_str, index) => {
          const modulo_num = index + 1;
          const mapeamento = [];

          try {
            const range = XLSX.utils.decode_range(range_str);
            const canal_info = POSICOES_CANAIS[range_str];

            if (modulo_num <= 5) {
              console.log(
                `Processando módulo ${modulo_num} (${range_str}), range:`,
                range
              );
            }

            // Verifica cada célula no range do módulo
            for (let row = range.s.r; row <= range.e.r; row++) {
              for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({
                  r: row,
                  c: col,
                });
                const cell = worksheet[cellAddress];

                if (cell && cell.v) {
                  // CORREÇÃO: Usar cell.v em vez de cell.value
                  const cellValue = String(cell.v);
                  const texto_celula = cellValue.trim().toUpperCase();

                  if (modulo_num <= 5) {
                    // console.log(`  Célula ${cellAddress}: "${cellValue}"`); // Comentado
                  }

                  // Procura por códigos de cores
                  for (const [codigo, nome_cor] of Object.entries(COLOR_MAP)) {
                    if (texto_celula === codigo) {
                      // CORREÇÃO: Usar igualdade estrita para evitar falsos positivos (ex: "BR" em "SOBRADO")
                      console.log(
                        `✓ Cor ${codigo} (${nome_cor}) encontrada no módulo ${modulo_num} em ${cellAddress}`
                      );

                      // Calcula posição do canal baseado no offset
                      let coord_canal_final = null;
                      let numero_canal = null;

                      if (canal_info) {
                        const [row_offset, col_offset] = canal_info;
                        const nova_row = row + row_offset;
                        const nova_col = col + col_offset;

                        if (nova_row >= 0 && nova_col >= 0) {
                          coord_canal_final = XLSX.utils.encode_cell({
                            r: nova_row,
                            c: nova_col,
                          });
                          const canal_cell = worksheet[coord_canal_final];

                          if (canal_cell && canal_cell.v) {
                            // CORREÇÃO: Usar cell.v em vez de cell.value
                            console.log(
                              `  Canal em ${coord_canal_final}: "${canal_cell.v}"`
                            );
                            if (!isNaN(canal_cell.v)) {
                              numero_canal = parseInt(canal_cell.v);
                              console.log(`  ✓ Canal ${numero_canal} válido`);
                            } else {
                              console.log(
                                `  ⚠ Valor do canal não é numérico: "${canal_cell.v}"`
                              );
                            }
                          } else {
                            console.log(
                              `  ⚠ Célula de canal ${coord_canal_final} vazia ou inexistente`
                            );
                          }
                        }
                      }

                      mapeamento.push({
                        cor: nome_cor,
                        canal: numero_canal,
                        canal_encontrado_em: coord_canal_final,
                      });
                    }
                  }
                }
              }
            }
          } catch (e) {
            console.warn(
              `Erro ao processar módulo ${modulo_num} (${range_str}):`,
              e
            );
          }

          mapeamento_cores_modulos.push({
            modulo: modulo_num,
            posicao: range_str,
            mapeamento: mapeamento,
          });

          if (mapeamento.length > 0) {
            console.log(
              `Módulo ${modulo_num}: ${mapeamento.length} cores encontradas`
            );
          }
        });

        // Conta módulos ativos (que têm pelo menos um mapeamento)
        const modulos_ativos = mapeamento_cores_modulos.filter(
          (mod) =>
            mod.mapeamento &&
            mod.mapeamento.length > 0 &&
            mod.mapeamento.some((map) => map.canal !== null && map.canal !== 0)
        ).length;

        console.log(`Total de módulos ativos encontrados: ${modulos_ativos}`);
        console.log("Informações do produto extraídas:", productInfo);

        // Converte para formato compatível com a detecção de layout
        const layout_para_deteccao = [];
        mapeamento_cores_modulos.forEach((modulo) => {
          if (modulo.mapeamento && modulo.mapeamento.length > 0) {
            // Pega a primeira cor encontrada para cada módulo
            const primeira_cor = modulo.mapeamento.find(
              (map) => map.cor && map.canal
            );
            if (primeira_cor) {
              layout_para_deteccao.push({
                module: modulo.modulo,
                color: primeira_cor.cor,
              });
              console.log(
                `Módulo ${modulo.modulo}: ${primeira_cor.cor} (canal ${primeira_cor.canal})`
              );
            }
          }
        });

        console.log(
          `Layout para detecção: ${layout_para_deteccao.length} módulos`
        );

        // Estrutura final igual ao Python
        const resultado_final = {
          codigo_firmware: productInfo.codigo_firmware,
          produto: productInfo.produto,
          modelo: productInfo.tecnologia, // CORREÇÃO: Usar o campo 'tecnologia' que foi preenchido
          cores_leds: productInfo.cores_leds,
          funcoes_extras: productInfo.funcoes_extras,
          cores_tampas: productInfo.cores_tampas,
          tamanho_polegadas: productInfo.tamanho_polegadas,
          quantidade_modulos: productInfo.quantidade_modulos,
          numero_leds: productInfo.numero_leds,
          cores_leds_lista: productInfo.cores_leds_lista,
          mapeamento_cores_modulos: mapeamento_cores_modulos,
          quantidade_modulos_original: productInfo.quantidade_modulos,
          quantidade_modulos_ativos: modulos_ativos,
        };

        console.log(
          "Processamento JavaScript (formato Python):",
          resultado_final
        );

        if (modulos_ativos === 0) {
          console.log("⚠ NENHUM MÓDULO ATIVO ENCONTRADO!");
          console.log("Verifique se:");
          console.log("1. As posições dos módulos estão corretas");
          console.log(
            "2. Os códigos de cores (VM, AZ, BR, AB) estão na planilha"
          );
          console.log("3. A planilha está no formato esperado");
        }

        resolve({
          layout: layout_para_deteccao,
          pythonData: resultado_final,
        });
      } catch (error) {
        console.error("Erro no processamento JavaScript:", error);
        reject(error);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

async function saveOrUpdateLayout() {
  if (!currentLayout || !activeBlockLayout) {
    alert("Nenhum layout ativo para salvar.");
    return;
  }

  const payload = {
    layout_id: currentLayout.id,
    nome: currentLayout.name,
    imagem: currentLayout.image,
    layout_data: activeBlockLayout,
  };

  try {
    const response = await fetch("salvar_layout.php", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Erro do servidor: ${response.status}. Detalhes: ${errorText}`
      );
    }

    const result = await response.json();

    if (result.success) {
      alert(
        `Sucesso! O layout foi salvo. A página será recarregada para aplicar as mudanças.`
      );
      localStorage.setItem("lastSavedLayoutId", payload.layout_id);
      location.reload();
    } else {
      throw new Error(
        result.error || "Erro desconhecido retornado pelo servidor."
      );
    }
  } catch (error) {
    console.error("ERRO AO SALVAR LAYOUT:", error);
    alert(`FALHA! Ocorreu um erro: ${error.message}`);
  }
}

function exportLayoutCoordinates() {
  if (!currentLayout || !activeBlockLayout || activeBlockLayout.length === 0) {
    alert("Não há layout ativo para exportar!");
    return;
  }

  const layoutName = `${currentLayout.id}_BlockLayout`;
  const layoutDataString = JSON.stringify(activeBlockLayout, null, 4);
  const codeToCopy = `const ${layoutName} = ${layoutDataString};`;

  navigator.clipboard
    .writeText(codeToCopy)
    .then(() => {
      alert(
        `Coordenadas do layout "${currentLayout.name}" copiadas para a área de transferência!`
      );
    })
    .catch((err) => {
      console.error("Falha ao copiar: ", err);
      alert(
        "Não foi possível copiar o código. Verifique o console para mais detalhes."
      );
    });
}

function initializeEventListeners() {
  if (openModalBtn)
    openModalBtn.addEventListener(
      "click",
      () => (configModal.style.display = "flex")
    );
  if (closeModalBtn)
    closeModalBtn.addEventListener(
      "click",
      () => (configModal.style.display = "none")
    );
  if (saveAndCloseModalBtn)
    saveAndCloseModalBtn.addEventListener("click", () => {
      configModal.style.display = "none";
    });
  if (openAdvancedSettingsBtn)
    openAdvancedSettingsBtn.addEventListener("click", () => {
      clearSelection();
      advancedSettingsModal.style.display = "flex";
    });
  if (closeAdvancedSettingsBtn)
    closeAdvancedSettingsBtn.addEventListener(
      "click",
      () => (advancedSettingsModal.style.display = "none")
    );
  if (saveAdvancedSettingsBtn)
    saveAdvancedSettingsBtn.addEventListener("click", saveAdvancedSettings);
  if (openLayoutPanelBtn)
    openLayoutPanelBtn.addEventListener("click", () => {
      populateLayoutPanel();
      layoutPanelModal.style.display = "flex";
    });
  if (closeLayoutPanelBtn)
    closeLayoutPanelBtn.addEventListener(
      "click",
      () => (layoutPanelModal.style.display = "none")
    );
  if (tabButtons)
    tabButtons.forEach((button) => {
      button.addEventListener("click", () => {
        tabButtons.forEach((btn) => btn.classList.remove("active"));
        tabContents.forEach((content) => content.classList.remove("active"));
        button.classList.add("active");
        document.getElementById(button.dataset.tab).classList.add("active");
      });
    });
  if (jumpTabelaCheckboxes)
    jumpTabelaCheckboxes.forEach((checkbox) => {
      checkbox.addEventListener("change", updateJumpTabelaVisuals);
    });
  if (specialFunctionCells)
    specialFunctionCells.forEach((cell) => {
      cell.addEventListener("input", () => updateSpecialFunctionVisuals(cell));
    });
  if (sheetSelectorContainer)
    sheetSelectorContainer.addEventListener("click", (e) => {
      if (e.target.matches(".sheet-selector-btn")) {
        switchSheet(e.target.dataset.sheet);
      }
    });
  if (limpaBtn) limpaBtn.addEventListener("click", clearCurrentSheet);
  if (gravarBtn) gravarBtn.addEventListener("click", saveSimulationToDatabase);
  if (restaurarBtn)
    restaurarBtn.addEventListener("click", () => {
      alert("Lembrar de ativar essa função");
    });
  if (executeMacroBtn) executeMacroBtn.addEventListener("click", executeMacro);

  // Listeners dos botões de layout
  if (saveLayoutBtn)
    saveLayoutBtn.addEventListener("click", saveOrUpdateLayout);
  const exportLayoutBtn = document.getElementById("export-layout-btn");
  if (exportLayoutBtn)
    exportLayoutBtn.addEventListener("click", exportLayoutCoordinates);

  const importBtn = document.getElementById("import-txt-btn");
  if (importBtn) {
    const fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = ".txt,text/plain";
    fileInput.style.display = "none";
    document.body.appendChild(fileInput);
    importBtn.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", handleFileImport);
  }

  const extractExcelBtn = document.getElementById("extract-excel-layout-btn");
  const excelLayoutInput = document.getElementById("excel-layout-input");

  if (extractExcelBtn && excelLayoutInput) {
    extractExcelBtn.addEventListener("click", () => {
      excelLayoutInput.click();
    });
    excelLayoutInput.addEventListener("change", async (event) => {
      const file = event.target.files[0];
      if (!file) return;
      extractExcelBtn.disabled = true;
      extractExcelBtn.innerHTML =
        '<i class="fa fa-spinner fa-spin"></i> Processando...';

      try {
        // Mantemos o try/catch para lidar com erros do próprio JS
        // ALTERAÇÃO: Chamando a função JavaScript diretamente, pulando a tentativa do Python.
        const extractedResult = await processExcelLayoutJS(file); // Retorna { layout: [...], pythonData: {...} }

        if (extractedResult.layout.length > 0) {
          // Detecta layout automaticamente baseado nos dados extraídos
          const detectedLayoutId = detectLayoutFromData(
            extractedResult.pythonData, // CORREÇÃO: Usar a chave correta 'pythonData'
            extractedResult
          );

          if (detectedLayoutId) {
            // ✅ CORREÇÃO: A lógica de aplicar cores foi movida para DENTRO da validação de sucesso.
            // 1. Limpa as cores de colunas anteriores
            columnColors = {};
            for (let i = 0; i < 24; i++) {
              applyColumnColor(i.toString(), "clear");
            }

            // Mapa de tradução de cores
            const colorTranslationMap = {
              vermelho: "red",
              azul: "blue",
              branco: "white",
              ambar: "yellow",
              verde: "green",
            };

            // 2. Itera sobre o mapeamento extraído do Excel para aplicar as novas cores
            const mapeamentoCompleto =
              extractedResult.pythonData.mapeamento_cores_modulos;
            mapeamentoCompleto.forEach((modulo) => {
              modulo.mapeamento.forEach((map) => {
                if (map.canal !== null && map.canal >= 1 && map.canal <= 24) {
                  const colIndex = map.canal - 1;
                  if (!columnColors[colIndex]) {
                    const translatedColor =
                      colorTranslationMap[map.cor.toLowerCase()];
                    if (translatedColor) {
                      applyColumnColor(colIndex.toString(), translatedColor);
                    }
                  }
                }
              });
            });
            console.log(
              "Cores das colunas aplicadas com base nos canais do Excel."
            );

            console.log(`Layout base detectado: ${detectedLayoutId}`);
            const baseLayoutTemplate = allAvailableLayouts.find(
              (l) => l.id === detectedLayoutId
            );

            if (!baseLayoutTemplate) {
              throw new Error(
                `Layout base '${detectedLayoutId}' não encontrado na lista de layouts disponíveis.`
              );
            }

            // --- INÍCIO DA LÓGICA DE CONSTRUÇÃO DINÂMICA ---
            // 1. Cria os placeholders de fundo primeiro, baseados no template da barra.
            const placeholderLayout = baseLayoutTemplate.blocks.map(
              (templateBlock) => ({
                ...templateBlock,
                type: "placeholder", // Identifica como um bloco de fundo
              })
            );
            const dynamicBlockLayout = [];
            // CORREÇÃO: Usar a chave correta 'pythonData'
            const mapeamentoDinamico =
              extractedResult.pythonData.mapeamento_cores_modulos;
            // ✅ NOVO: Filtra apenas os módulos que estão ativos (têm canal) no Excel.
            const modulosAtivosDoExcel = mapeamentoDinamico.filter(
              (m) =>
                m.mapeamento &&
                m.mapeamento.length > 0 &&
                m.mapeamento.some((map) => map.canal !== null && map.canal > 0)
            );

            // Itera sobre as posições FÍSICAS do template da barra (os blocos visuais)
            baseLayoutTemplate.blocks.forEach((templateBlock) => {
              const physicalPositionId = templateBlock.id;
              const physicalIndex = parseInt(physicalPositionId, 10) - 1;

              // ✅ LÓGICA DINÂMICA: Associa o bloco físico ao módulo do Excel pela ordem.
              // O bloco físico '1' pega o primeiro módulo ativo, o bloco '2' o segundo, etc.
              if (
                physicalIndex >= 0 &&
                physicalIndex < modulosAtivosDoExcel.length
              ) {
                const excelModuleData = modulosAtivosDoExcel[physicalIndex];

                // Para cada canal encontrado neste módulo, cria um bloco visual
                excelModuleData.mapeamento.forEach((map) => {
                  if (map.canal !== null && map.canal > 0) {
                    const newBlock = {
                      ...templateBlock, // Usa a posição física do template
                      text: String(map.canal), // Exibe o número do canal
                    };
                    dynamicBlockLayout.push(newBlock);
                  }
                });
              }
            });

            console.log(
              `Layout dinâmico gerado com ${dynamicBlockLayout.length} blocos visuais.`
            );

            // ATUALIZA O ESTADO DA APLICAÇÃO COM O LAYOUT DINÂMICO
            currentLayout = baseLayoutTemplate; // Mantém o nome e a imagem do layout base
            // Combina os placeholders com os blocos dinâmicos para renderização
            activeBlockLayout = [...placeholderLayout, ...dynamicBlockLayout];

            // ATUALIZA A INTERFACE
            const placeholder = document.getElementById(
              "simulation-placeholder"
            );
            if (placeholder) placeholder.style.display = "none";
            currentLayoutNameSpan.textContent =
              currentLayout.name + " (Dinâmico)";
            document.getElementById("siren-model-image").src =
              currentLayout.image;
            initializeBlocks(activeBlockLayout); // Renderiza os blocos dinâmicos
            stopSimulation();
            // --- FIM DA LÓGICA DE CONSTRUÇÃO DINÂMICA ---
          } else {
            alert(
              "Não foi possível detectar um layout base para os dados extraídos."
            );
          }

          console.log(`Layout dinâmico gerado com sucesso a partir do Excel!`);
        } else {
          alert(
            "Não foi possível extrair dados do arquivo Excel. Verifique o formato do arquivo."
          );
        }
      } catch (error) {
        console.error("Erro ao extrair layout do Excel:", error);
        alert(`Falha ao processar o arquivo Excel. Detalhes: ${error.message}`);
      } finally {
        extractExcelBtn.disabled = false;
        extractExcelBtn.innerHTML =
          '<i class="fa fa-file-excel"></i> Extrair Layout Excel';
        excelLayoutInput.value = "";
      }
    });
  }

  const exportGifBtn = document.getElementById("export-gif-btn");
  if (exportGifBtn) {
    exportGifBtn.addEventListener("click", exportAllSimulationsAsGifs);
  }

  if (editLayoutBtn && simulationArea && saveLayoutBtn) {
    editLayoutBtn.addEventListener("click", () => {
      // Se a simulação estiver rodando, pare-a antes de entrar no modo de edição.
      if (isSimulationRunning) {
        stopSimulation();
      }

      isEditMode = !isEditMode;
      simulationArea.classList.toggle("edit-mode", isEditMode);

      const exportLayoutBtn = document.getElementById("export-layout-btn");

      if (isEditMode) {
        editLayoutBtn.innerHTML = '<i class="fa fa-check"></i> Concluir Edição';
        saveLayoutBtn.style.display = "inline-block";
        exportLayoutBtn.style.display = "inline-block";
        stopSimulation();
      } else {
        editLayoutBtn.innerHTML = '<i class="fa fa-pencil"></i> Editar Layout';
        editLayoutBtn.style.backgroundColor = "#3498db";
        saveLayoutBtn.style.display = "none";
        exportLayoutBtn.style.display = "none";
        if (currentLayout) {
          const blockElements =
            blocksContainer.querySelectorAll(".block-container");
          blockElements.forEach((blockEl, index) => {
            if (activeBlockLayout[index]) {
              activeBlockLayout[index].top = blockEl.style.top;
              activeBlockLayout[index].left = blockEl.style.left;
              activeBlockLayout[index].width = parseFloat(blockEl.style.width);
              activeBlockLayout[index].height = parseFloat(
                blockEl.style.height
              );
              const rotation =
                blockEl.style.transform.match(/rotate\((.+)deg\)/);
              activeBlockLayout[index].rotate = rotation
                ? parseFloat(rotation[1])
                : 0;
            }
          });
        }
        if (activeBlock) {
          activeBlock.classList.remove("selected");
          activeBlock = null;
        }
      }
    });
  }

  const hideContextMenu = () => {
    if (colorContextMenu) colorContextMenu.style.display = "none";
    window.removeEventListener("click", hideContextMenu);
  };

  const showColorMenu = (e) => {
    const headerCell = e.target.closest("th.column-header[data-col-index]");
    if (headerCell) {
      e.preventDefault();
      if (colorContextMenu) {
        colorContextMenu.style.display = "block";
        colorContextMenu.style.top = `${e.clientY}px`;
        colorContextMenu.style.left = `${e.clientX}px`;
        colorContextMenu.dataset.editingCol = headerCell.dataset.colIndex;
        setTimeout(() => {
          window.addEventListener("click", hideContextMenu, {
            once: true,
          });
        }, 0);
      }
    }
  };

  if (macroHeader) macroHeader.addEventListener("contextmenu", showColorMenu);

  if (colorContextMenu)
    colorContextMenu.addEventListener("click", (e) => {
      e.stopPropagation();
      const target = e.target.closest("li[data-color]");
      if (target) {
        const colIndex = colorContextMenu.dataset.editingCol;
        const color = target.dataset.color;
        applyColumnColor(colIndex, color);
        hideContextMenu();
      }
    });

  if (blocksContainer)
    blocksContainer.addEventListener("mousedown", (e) => {
      // Impede a seleção de placeholders
      if (e.target.closest(".block-placeholder")) {
        return;
      }

      if (!isEditMode) return;
      const target = e.target;
      if (target.isContentEditable) {
        return;
      }
      const blockElement = target.closest(".block-container");
      if (!blockElement) {
        if (activeBlock) {
          activeBlock.classList.remove("selected");
          activeBlock = null;
        }
        return;
      }
      e.preventDefault();
      if (activeBlock && activeBlock !== blockElement) {
        activeBlock.classList.remove("selected");
      }
      activeBlock = blockElement;
      activeBlock.classList.add("selected");
      const initialMouseX = e.clientX;
      const initialMouseY = e.clientY;
      const blockRect = activeBlock.getBoundingClientRect();
      if (target.classList.contains("resize-handle")) {
        action = {
          type: "resize",
          initialWidth: blockRect.width,
          initialHeight: blockRect.height,
          initialMouseX,
          initialMouseY,
        };
      } else if (target.classList.contains("rotate-handle")) {
        const centerX = blockRect.left + blockRect.width / 2;
        const centerY = blockRect.top + blockRect.height / 2;
        action = {
          type: "rotate",
          centerX,
          centerY,
        };
      } else {
        const initialTop = activeBlock.offsetTop;
        const initialLeft = activeBlock.offsetLeft;
        action = {
          type: "move",
          initialTop,
          initialLeft,
          initialMouseX,
          initialMouseY,
        };
      }
      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    });
}

function onMouseMove(e) {
  if (!action || !activeBlock) return;
  // Impede a movimentação de placeholders
  if (activeBlock.classList.contains("block-placeholder")) {
    return;
  }

  const blockData = activeBlockLayout[activeBlock.dataset.index];
  const containerRect = simulationArea.getBoundingClientRect();
  if (action.type === "move") {
    const dx = e.clientX - action.initialMouseX;
    const dy = e.clientY - action.initialMouseY;
    activeBlock.style.left = `${
      ((action.initialLeft + dx) / containerRect.width) * 100
    }%`;
    activeBlock.style.top = `${
      ((action.initialTop + dy) / containerRect.height) * 100
    }%`;
    blockData.left = activeBlock.style.left;
    blockData.top = activeBlock.style.top;
  } else if (action.type === "resize") {
    const dx = e.clientX - action.initialMouseX;
    const dy = e.clientY - action.initialMouseY;
    let newWidth = action.initialWidth + dx;
    let newHeight = action.initialHeight + dy;
    if (newWidth > 20) {
      activeBlock.style.width = `${newWidth}px`;
      blockData.width = newWidth;
    }
    if (newHeight > 10) {
      activeBlock.style.height = `${newHeight}px`;
      blockData.height = newHeight;
    }
  } else if (action.type === "rotate") {
    const angleRad = Math.atan2(
      e.clientY - action.centerY,
      e.clientX - action.centerX
    );
    let angleDeg = angleRad * (180 / Math.PI) + 90;
    activeBlock.style.transform = `rotate(${angleDeg}deg)`;
    blockData.rotate = angleDeg;
  }
}

function onMouseUp() {
  action = null;
  document.removeEventListener("mousemove", onMouseMove);
  document.removeEventListener("mouseup", onMouseUp);
}

function clearSelection() {
  selection.forEach((cell) => cell.classList.remove("cell-selected"));
  selection.clear();
}

function toggleSpecialFunction(funcName) {
  if (activeSpecialFunctions[funcName]) {
    delete activeSpecialFunctions[funcName];
  } else {
    const onList = savedSpecialFunctionConfigs[funcName + "_on"] || [];
    const macroData = allMacroData[activeSheet];
    const newColors = {};
    onList.forEach((blockNumber) => {
      let foundColor = null;
      for (const row of macroData) {
        const colorValue = row.values[blockNumber - 1];
        if (colorValue && ["1", "2", "3", "4"].includes(colorValue)) {
          foundColor = parseInt(colorValue, 10);
          break;
        }
      }
      newColors[blockNumber] = foundColor || 3;
    });
    activeSpecialFunctions[funcName] = {
      colors: newColors,
    };
  }
  if (!isSimulationRunning) {
    renderLights();
  }
  updateSpecialFunctionUI();
}

function updateSpecialFunctionUI() {
  document
    .querySelectorAll(".control-btn")
    .forEach((btn) => btn.classList.remove("active"));
  for (const funcName in activeSpecialFunctions) {
    let btnId;
    if (funcName === "takedown") btnId = "btn-td";
    if (funcName === "alley_left") btnId = "btn-le";
    if (funcName === "alley_right") btnId = "btn-ld";
    if (btnId) {
      const btn = document.getElementById(btnId);
      if (btn) btn.classList.add("active");
    }
  }
}

function triggerMacro(sheetName) {
  if (isSimulationRunning && activeSheet === sheetName) {
    stopSimulation();
    return;
  }
  if (!sheetNames.includes(sheetName)) {
    console.warn(
      `Tentativa de executar macro para uma aba inexistente: ${sheetName}`
    );
    return;
  }
  stopSimulation();
  setTimeout(() => {
    switchSheet(sheetName);
    executeMacro();
  }, 100);
}

function executeMacro() {
  if (isSimulationRunning) {
    stopSimulation();
    return;
  }
  allMacroData[activeSheet] = getCurrentTableData();
  const rawData = allMacroData[activeSheet];

  const macroSteps = [];
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const msValue = parseInt(row.ms, 10);

    if (isNaN(msValue) || row.ms.trim() === "") {
      break;
    }

    macroSteps.push({
      ...row,
      originalIndex: i,
    });

    if (msValue === 0) {
      break;
    }
  }

  if (macroSteps.length === 0) {
    alert(
      "Nenhum passo válido encontrado na planilha para iniciar a simulação!"
    );
    return;
  }

  startSimulation(macroSteps);
}

function startSimulation(macroData) {
  isSimulationRunning = true;
  executeMacroBtn.textContent = "Parar Macro";
  executeMacroBtn.style.backgroundColor = "#dc3545";
  let currentStep = 0;

  function runStep() {
    if (!isSimulationRunning) return;

    const step = macroData[currentStep];
    if (!step) {
      currentStep = 0;
      runStep();
      return;
    }

    const stepTime = parseInt(step.ms, 10);
    if (stepTime === 0) {
      currentStep = 0;
      runStep();
      return;
    }

    const values = step.values.map((v) => parseInt(v, 10) || 0);
    renderLights(values);
    highlightCurrentRow(step.originalIndex);
    currentStep++;

    if (currentStep >= macroData.length) {
      currentStep = 0;
    }

    simulationStopper = setTimeout(runStep, stepTime);
  }
  runStep();
}

function stopSimulation() {
  isSimulationRunning = false;
  if (executeMacroBtn) {
    executeMacroBtn.textContent = "Executar Macro";
    executeMacroBtn.style.backgroundColor = "#f39c12";
  }
  clearTimeout(simulationStopper);
  simulationStopper = null;
  renderLights();
  document
    .querySelectorAll("#macro-body tr.selected-row")
    .forEach((tr) => tr.classList.remove("selected-row"));
}

function renderLights(macroValues = []) {
  const finalOnColors = {};
  const finalOffList = [];
  const colorClassMap = {
    red: "layout-color-red",
    green: "layout-color-green",
    blue: "layout-color-blue",
    white: "layout-color-white",
    yellow: "layout-color-yellow",
  };

  for (const funcName in activeSpecialFunctions) {
    const offConfig = savedSpecialFunctionConfigs[funcName] || [];
    offConfig.forEach((blockNum) => finalOffList.push(blockNum));
    const onConfig = activeSpecialFunctions[funcName].colors || {};
    for (const blockNum in onConfig) {
      finalOnColors[blockNum] = onConfig[blockNum];
    }
  }
  const cutHornList = savedSpecialFunctionConfigs.cut_horn_light || [];
  cutHornList.forEach((blockNum) => finalOffList.push(blockNum));
  let brightness = 1.0;
  if (savedJumpTabelaSettings["jt2_bit7"]) brightness = 0.75;
  if (savedJumpTabelaSettings["jt2_bit6"]) brightness = 0.5;

  document.querySelectorAll("[data-block-num]").forEach((partElement) => {
    const blockNumber = parseInt(partElement.dataset.blockNum, 10);
    if (isNaN(blockNumber)) return;

    const blockContainer = partElement.closest(".block-container");
    const columnIndex = blockNumber - 1;

    let finalColorValue = macroValues[columnIndex] || 0;
    if (finalOnColors.hasOwnProperty(blockNumber)) {
      finalColorValue = finalOnColors[blockNumber];
    }
    if (finalOffList.includes(blockNumber)) {
      finalColorValue = 0;
    }

    partElement.className = partElement.className
      .replace(/layout-color-\w+/g, "")
      .trim();
    partElement.classList.add("layout-color-off");

    // Durante a simulação (macroValues tem itens), esconde por padrão.
    // Se a simulação estiver parada (macroValues está vazio), mostra todos.
    if (blockContainer) {
      blockContainer.style.opacity = macroValues.length > 0 ? "0" : "1.0";
    }

    if (finalColorValue > 0) {
      if (blockContainer) {
        // Torna o bloco visível apenas se ele estiver ativo neste passo.
        // A opacidade final será o brilho definido (ex: 1.0, 0.75, 0.5).
        blockContainer.style.opacity = brightness;
      }
      const persistentColor = columnColors[columnIndex];
      let colorToApply = "";

      if (persistentColor && colorClassMap[persistentColor]) {
        colorToApply = colorClassMap[persistentColor];
      } else {
        // Se não houver cor de coluna, usa vermelho como padrão para o valor "1".
        if (finalColorValue === 1) colorToApply = "layout-color-red";
      }

      if (colorToApply) {
        partElement.classList.remove("layout-color-off");
        partElement.classList.add(colorToApply);
      }
    }
  });
}

function previewRow(rowIndex) {
  if (isSimulationRunning) return;
  const tableData = getCurrentTableData();
  if (rowIndex >= 0 && rowIndex < tableData.length) {
    const rowData = tableData[rowIndex];
    const values = rowData.values.map((v) => parseInt(v, 10) || 0);
    renderLights(values);
  }
}

function highlightCurrentRow(rowIndex) {
  document
    .querySelectorAll("#macro-body tr.selected-row")
    .forEach((tr) => tr.classList.remove("selected-row"));
  const targetRow = macroTableBody.rows[rowIndex];
  if (targetRow) {
    targetRow.classList.add("selected-row");
    targetRow.scrollIntoView({
      behavior: "smooth",
      block: "nearest",
    });
    previewRow(rowIndex);
  }
}

function getCurrentTableData() {
  return Array.from(macroTableBody.rows).map((row) => {
    const cells = Array.from(row.cells);
    return {
      row: cells[2].textContent.trim(),
      ms: cells[1].textContent.trim(),
      values: cells.slice(3).map((cell) => cell.textContent.trim()),
    };
  });
}

function switchSheet(newSheetName, preventSave = false) {
  if (activeSheet === newSheetName) return;
  if (!preventSave) {
    allMacroData[activeSheet] = getCurrentTableData();
  }
  activeSheet = newSheetName;
  renderTable(allMacroData[activeSheet]);
  updateActiveSheetUI();
  updateConsumptionDisplay();
}

function updateActiveSheetUI() {
  document.querySelectorAll(".sheet-selector-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.sheet === activeSheet);
  });
}

function clearCurrentSheet() {
  if (
    confirm(
      `Tem certeza que deseja limpar todos os dados da planilha "${activeSheet}"? Esta ação pode ser desfeita com Ctrl+Z.`
    )
  ) {
    pushStateToUndoHistory();
    allMacroData[activeSheet] = JSON.parse(JSON.stringify(baseSheetStructure));
    columnColors = {};
    renderTable(allMacroData[activeSheet]);
    updateConsumptionDisplay();
  }
}

function saveAdvancedSettings() {
  const currentModalConfigs = readCurrentModalState();
  savedSpecialFunctionConfigs = currentModalConfigs.specialFunctions;
  savedJumpTabelaSettings = currentModalConfigs.jumpTabela;
  renderAllActiveSettings();
  alert("Definições avançadas salvas!");
  advancedSettingsModal.style.display = "none";
}

function readCurrentModalState() {
  const specialFunctions = {};
  document
    .querySelectorAll("#tab-special-functions .function-table-container")
    .forEach((container) => {
      const funcName =
        container.querySelector("td[data-function]").dataset.function;
      const cells = container.querySelectorAll("td[data-block]");
      const currentValues = [];
      cells.forEach((cell) => {
        if (cell.textContent.trim() === "1") {
          currentValues.push(parseInt(cell.dataset.block, 10));
        }
      });
      if (currentValues.length > 0) {
        specialFunctions[funcName] = currentValues;
      }
    });

  const jumpTabela = {};
  jumpTabelaCheckboxes.forEach((checkbox) => {
    const key = `jt${checkbox.dataset.jumptabela}_bit${checkbox.dataset.bit}`;
    jumpTabela[key] = checkbox.checked;
  });

  return {
    specialFunctions,
    jumpTabela,
  };
}

function updateAllVisuals() {
  updateJumpTabelaVisuals();
  specialFunctionCells.forEach((cell) => updateSpecialFunctionVisuals(cell));
}

function updateJumpTabelaVisuals() {
  jumpTabelaCheckboxes.forEach((checkbox) => {
    checkbox
      .closest(".checkbox-item")
      .classList.toggle("active-setting", checkbox.checked);
  });
}

function updateSpecialFunctionVisuals(cell) {
  cell.classList.toggle(
    "cell-active-function",
    cell.textContent.trim() === "1"
  );
}

function renderAllActiveSettings() {
  renderActiveSpecialFunctions();
  renderActiveJumpTabelaSettings();
}

function renderActiveSpecialFunctions() {
  specialFunctionsStatusContainer.innerHTML = "";
  const functionNamesMap = {
    takedown: "Take Down (Apagar)",
    takedown_on: "Take Down (Acender)",
    alley_right: "Alley-Right (LD - Apagar)",
    alley_right_on: "Alley-Right (LD - Acender)",
    alley_left: "Alley-Left (LE - Apagar)",
    alley_left_on: "Alley-Left (LE - Acender)",
    backlight_off: "Backlight (Apagar)",
    backlight_on: "Backlight (Acender)",
    cut_front_off: "Cut Front (Apagar)",
    cut_front_on: "Cut Front (Acender)",
    cut_rear_off: "Cut Rear (Apagar)",
    cut_rear_on: "Cut Rear (Acender)",
    cut_dt_off: "Cut DT (Apagar)",
    cut_dt_on: "Cut DT (Acender)",
    cut_horn_light: "Cut Horn Light",
  };

  const activeFunctions = [];
  for (const funcKey in savedSpecialFunctionConfigs) {
    const configArray = savedSpecialFunctionConfigs[funcKey];
    if (configArray && configArray.length > 0) {
      const funcName = functionNamesMap[funcKey] || funcKey;
      const blocksText = configArray.join(", ");
      activeFunctions.push(
        `<span class="func-name">${funcName}:</span> Bloco(s) ${blocksText}`
      );
    }
  }

  if (activeFunctions.length > 0) {
    const ul = document.createElement("ul");
    activeFunctions.forEach((text) => {
      const li = document.createElement("li");
      li.innerHTML = text;
      ul.appendChild(li);
    });
    specialFunctionsStatusContainer.appendChild(ul);
  } else {
    specialFunctionsStatusContainer.innerHTML =
      '<p class="no-settings">Nenhuma função especial ativa.</p>';
  }
}

function renderActiveJumpTabelaSettings() {
  jumpTabelaStatusContainer.innerHTML = "";
  const activeSettings = [];
  for (const key in savedJumpTabelaSettings) {
    if (savedJumpTabelaSettings[key]) {
      const checkbox = document.querySelector(
        `[data-jumptabela="${key[2]}"][data-bit="${key[8]}"]`
      );
      if (checkbox) {
        activeSettings.push(checkbox.parentElement.textContent.trim());
      }
    }
  }

  if (activeSettings.length > 0) {
    const ul = document.createElement("ul");
    activeSettings.forEach((text) => {
      const li = document.createElement("li");
      li.textContent = text;
      ul.appendChild(li);
    });
    jumpTabelaStatusContainer.appendChild(ul);
  } else {
    jumpTabelaStatusContainer.innerHTML =
      '<p class="no-settings">Nenhuma configuração global ativa.</p>';
  }
}

function initializeInteractionHandlers() {
  let isMouseDown = false,
    startCell = null;
  if (macroTableBody)
    macroTableBody.addEventListener("mousedown", (e) => {
      const targetCell = e.target.closest('td[contenteditable="true"]');
      if (!targetCell) return;
      isMouseDown = true;
      startCell = targetCell;
      if (!e.shiftKey) {
        clearSelection();
        selection.add(targetCell);
        targetCell.classList.add("cell-selected");
      }
    });
  if (macroTableBody)
    macroTableBody.addEventListener("mousemove", (e) => {
      if (!isMouseDown) return;
      document.body.style.userSelect = "none";
      const endCell = e.target.closest('td[contenteditable="true"]');
      if (!endCell) return;
      clearSelection();
      const allRows = Array.from(macroTableBody.rows);
      const startRowIndex = startCell.parentElement.rowIndex;
      const endRowIndex = endCell.parentElement.rowIndex;
      const startCellIndex = startCell.cellIndex;
      const endCellIndex = endCell.cellIndex;
      const minRow = Math.min(startRowIndex, endRowIndex);
      const maxRow = Math.max(startRowIndex, endRowIndex);
      const minCol = Math.min(startCellIndex, endCellIndex);
      const maxCol = Math.max(startCellIndex, endCellIndex);
      for (let i = minRow; i <= maxRow; i++) {
        const row = allRows[i - 1];
        if (row) {
          for (let j = minCol; j <= maxCol; j++) {
            const cell = row.cells[j];
            if (cell && cell.hasAttribute("contenteditable")) {
              cell.classList.add("cell-selected");
              selection.add(cell);
            }
          }
        }
      }
    });
  document.addEventListener("mouseup", () => {
    isMouseDown = false;
    document.body.style.userSelect = "";
    startCell = null;
  });
  document.addEventListener("keydown", (e) => {
    if (isEditMode && activeBlock) {
      if (e.key === "Delete") {
        // Impede a exclusão de placeholders
        if (activeBlock.classList.contains("block-placeholder")) {
          return;
        }

        e.preventDefault();
        const indexToRemove = parseInt(activeBlock.dataset.index, 10);
        activeBlockLayout.splice(indexToRemove, 1);
        activeBlock = null;
        initializeBlocks(activeBlockLayout);
        return;
      }
      if (e.ctrlKey && e.key.toLowerCase() === "c") {
        e.preventDefault();
        // Impede a cópia de placeholders
        if (activeBlock.classList.contains("block-placeholder")) {
          return;
        }

        const indexToCopy = parseInt(activeBlock.dataset.index, 10);
        copiedBlockData = JSON.parse(
          JSON.stringify(activeBlockLayout[indexToCopy])
        );
        activeBlock.style.outline = "2px dashed green";
        setTimeout(() => {
          if (activeBlock) activeBlock.style.outline = "";
        }, 300);
        return;
      }
    }
    if (
      isEditMode &&
      e.ctrlKey &&
      e.key.toLowerCase() === "v" &&
      copiedBlockData
    ) {
      e.preventDefault();
      const newBlock = JSON.parse(JSON.stringify(copiedBlockData));
      const currentLeft = parseFloat(newBlock.left);
      const currentTop = parseFloat(newBlock.top);
      newBlock.left = `${currentLeft + 2}%`;
      newBlock.top = `${currentTop + 2}%`;
      activeBlockLayout.push(newBlock);
      initializeBlocks(activeBlockLayout);
      return;
    }

    if (e.ctrlKey) {
      switch (e.key.toLowerCase()) {
        case "c":
          if (selection.size > 0) {
            e.preventDefault();
            const cells = Array.from(selection);
            const minRow = Math.min(
              ...cells.map((c) => c.parentElement.rowIndex)
            );
            const minCol = Math.min(...cells.map((c) => c.cellIndex));
            clipboardData = cells.map((cell) => ({
              row: cell.parentElement.rowIndex - minRow,
              col: cell.cellIndex - minCol,
              value: cell.textContent,
            }));
            cells.forEach((c) => {
              c.style.outline = "1px dashed #0070c0";
            });
            setTimeout(
              () =>
                cells.forEach((c) => {
                  c.style.outline = "";
                }),
              200
            );
          }
          break;
        case "v":
          if (clipboardData && selection.size > 0) {
            e.preventDefault();
            pushStateToUndoHistory();
            let affectedCellsForSync = [];
            if (selection.size > 1) {
              const clipboardHeight =
                Math.max(...clipboardData.map((c) => c.row)) + 1;
              const clipboardWidth =
                Math.max(...clipboardData.map((c) => c.col)) + 1;
              const clipboardMap = new Map();
              clipboardData.forEach((c) =>
                clipboardMap.set(`${c.row},${c.col}`, c.value)
              );
              const selectedCells = Array.from(selection);
              affectedCellsForSync = selectedCells;
              const minSelectedRow = Math.min(
                ...selectedCells.map((c) => c.parentElement.rowIndex)
              );
              const minSelectedCol = Math.min(
                ...selectedCells.map((c) => c.cellIndex)
              );
              selectedCells.forEach((cell) => {
                const relativeRow =
                  cell.parentElement.rowIndex - minSelectedRow;
                const relativeCol = cell.cellIndex - minSelectedCol;
                const sourceRow = relativeRow % clipboardHeight;
                const sourceCol = relativeCol % clipboardWidth;
                const valueToPaste = clipboardMap.get(
                  `${sourceRow},${sourceCol}`
                );
                if (valueToPaste !== undefined) {
                  cell.textContent = valueToPaste;
                  updateCellColor(cell);
                }
              });
            } else if (selection.size === 1) {
              const targetCell = Array.from(selection)[0];
              const startRow = targetCell.parentElement.rowIndex;
              const startCol = targetCell.cellIndex;
              const allRows = Array.from(macroTableBody.rows);
              clipboardData.forEach((clip) => {
                const targetRowIndex = startRow + clip.row - 1;
                const targetColIndex = startCol + clip.col;
                if (allRows[targetRowIndex]) {
                  const finalCell =
                    allRows[targetRowIndex].cells[targetColIndex];
                  if (finalCell && finalCell.isContentEditable) {
                    finalCell.textContent = clip.value;
                    updateCellColor(finalCell);
                    affectedCellsForSync.push(finalCell);
                  }
                }
              });
            }
            affectedCellsForSync.forEach((cell) => {
              if (cell.classList.contains("ms25-column")) {
                const msCell = cell.nextElementSibling;
                if (msCell) syncTimeValues(cell, msCell, "from_ms25");
              } else if (cell.classList.contains("ms-column")) {
                const ms25Cell = cell.previousElementSibling;
                if (ms25Cell) syncTimeValues(ms25Cell, cell, "from_ms");
              }
            });
            updateConsumptionDisplay();
          }
          break;
        case "z":
          e.preventDefault();
          const undoHistoryForSheet = undoHistory[activeSheet];
          if (undoHistoryForSheet && undoHistoryForSheet.length > 0) {
            const previousState = undoHistoryForSheet.pop();
            redoHistory[activeSheet].push(getCurrentTableData());
            allMacroData[activeSheet] = previousState;
            renderTable(allMacroData[activeSheet]);
            updateConsumptionDisplay();
          }
          break;
        case "y":
          e.preventDefault();
          const redoHistoryForSheet = redoHistory[activeSheet];
          if (redoHistoryForSheet && redoHistoryForSheet.length > 0) {
            const nextState = redoHistoryForSheet.pop();
            undoHistory[activeSheet].push(getCurrentTableData());
            allMacroData[activeSheet] = nextState;
            renderTable(allMacroData[activeSheet]);
            updateConsumptionDisplay();
          }
          break;
      }
      return;
    }
    const activeEl = document.activeElement;
    if (activeEl && activeEl.closest(".function-table")) {
      const currentCell = activeEl;
      if (currentCell.tagName !== "TD" || !currentCell.isContentEditable)
        return;
      const currentRow = currentCell.parentElement;
      const currentCellIndex = currentCell.cellIndex;
      let nextCell;
      switch (e.key) {
        case "Enter":
        case "ArrowDown":
          e.preventDefault();
          let nextRow = currentRow.nextElementSibling;
          if (nextRow) {
            nextCell = nextRow.cells[currentCellIndex];
          } else {
            const currentTableContainer = currentRow.closest(
              ".function-table-container"
            );
            const nextTableContainer = currentTableContainer.nextElementSibling;
            if (
              nextTableContainer &&
              nextTableContainer.classList.contains("function-table-container")
            ) {
              nextCell =
                nextTableContainer.querySelector("tbody tr").cells[
                  currentCellIndex
                ];
            }
          }
          break;
        case "ArrowUp":
          e.preventDefault();
          let prevRow = currentRow.previousElementSibling;
          if (prevRow) {
            nextCell = prevRow.cells[currentCellIndex];
          } else {
            const currentTableContainer = currentRow.closest(
              ".function-table-container"
            );
            const prevTableContainer =
              currentTableContainer.previousElementSibling;
            if (
              prevTableContainer &&
              prevTableContainer.classList.contains("function-table-container")
            ) {
              nextCell = prevTableContainer.querySelector("tbody tr:last-child")
                .cells[currentCellIndex];
            }
          }
          break;
        case "ArrowLeft":
          if (window.getSelection().anchorOffset === 0) {
            e.preventDefault();
            nextCell = currentCell.previousElementSibling;
          }
          break;
        case "ArrowRight":
          if (
            window.getSelection().anchorOffset ===
              currentCell.textContent.length ||
            currentCell.textContent.length === 0
          ) {
            e.preventDefault();
            nextCell = currentCell.nextElementSibling;
          }
          break;
      }
      if (nextCell && nextCell.hasAttribute("contenteditable")) {
        nextCell.focus();
      }
      return;
    }
    if (selection.size > 0 || (activeEl && activeEl.closest(".macro-table"))) {
      if (e.key === "Delete" || e.key === "Backspace") {
        e.preventDefault();
        pushStateToUndoHistory();
        selection.forEach((cell) => {
          if (cell.isContentEditable) {
            cell.textContent = "";
            if (cell.classList.contains("editable-cell")) updateCellColor(cell);
          }
        });
        updateConsumptionDisplay();
        return;
      }
      if (selection.size > 1 && /[1-4]/.test(e.key)) {
        e.preventDefault();
        pushStateToUndoHistory();
        selection.forEach((cell) => {
          if (cell.classList.contains("editable-cell")) {
            cell.textContent = e.key;
            updateCellColor(cell);
          }
        });
        updateConsumptionDisplay();
        return;
      }
      if (activeEl.isContentEditable) {
        const currentCell = activeEl;
        const currentRow = currentCell.parentElement;
        const currentCellIndex = currentCell.cellIndex;
        let nextCell;
        switch (e.key) {
          case "Enter":
          case "ArrowDown":
            e.preventDefault();
            if (currentRow.nextElementSibling)
              nextCell = currentRow.nextElementSibling.cells[currentCellIndex];
            break;
          case "ArrowUp":
            e.preventDefault();
            if (currentRow.previousElementSibling)
              nextCell =
                currentRow.previousElementSibling.cells[currentCellIndex];
            break;
          case "ArrowLeft":
            if (window.getSelection().anchorOffset === 0) {
              e.preventDefault();
              nextCell = currentCell.previousElementSibling;
            }
            break;
          case "ArrowRight":
            if (
              window.getSelection().anchorOffset ===
                currentCell.textContent.length ||
              currentCell.textContent.length === 0
            ) {
              e.preventDefault();
              nextCell = currentCell.nextElementSibling;
            }
            break;
        }
        if (nextCell && nextCell.hasAttribute("contenteditable")) {
          nextCell.focus();
          clearSelection();
          selection.add(nextCell);
          nextCell.classList.add("cell-selected");
        }
      }
    }
  });
}

function handleFileImport(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const content = e.target.result;
    try {
      parseAndApplyTxtData(content);
      console.log("Arquivo TXT importado com sucesso!");
    } catch (error) {
      console.error("ERRO CRÍTICO ao processar o arquivo TXT:", error);
      alert(
        `Falha ao importar o arquivo. Verifique o formato e o console (F12).\nErro: ${error.message}`
      );
    }
    event.target.value = "";
  };
  reader.onerror = () => {
    alert("Não foi possível ler o arquivo.");
    event.target.value = "";
  };
  reader.readAsText(file);
}

function parseNumericValue(val) {
  if (typeof val !== "string") return val;
  const cleanVal = val.trim().toUpperCase();
  if (cleanVal.startsWith("0B")) return parseInt(cleanVal.substring(2), 2);
  if (cleanVal.startsWith("0X")) return parseInt(cleanVal.substring(2), 16);
  const num = parseInt(cleanVal, 10);
  return isNaN(num) ? null : num;
}

function bytesToTableValues(
  byteParaCols17a24,
  byteParaCols9a16,
  byteParaCols1a8,
  columnColors // Argumento adicionado
) {
  const values = Array(24).fill("");
  const bytes = [byteParaCols1a8, byteParaCols9a16, byteParaCols17a24];
  for (let byteIndex = 0; byteIndex < 3; byteIndex++) {
    const byteVal = bytes[byteIndex];
    if (byteVal === null) continue;
    for (let bitIndex = 0; bitIndex < 8; bitIndex++) {
      if ((byteVal >> bitIndex) & 1) {
        const columnIndex = byteIndex * 8 + bitIndex;
        if (columnIndex < 24) {
          // Lógica simplificada: se o bit está ativo, o valor é sempre "1".
          values[columnIndex] = "1";
        }
      }
    }
  }
  return values;
}

function parseAndApplyTxtData(content) {
  const populatedSheets = new Set();
  const tempData = {}; // Objeto temporário para guardar os dados lidos
  const paternsRegex = /Paterns\s*\[\s*[^\]]+\]\s*=\s*\{([\s\S]*?)\};/;
  const paternsMatch = content.match(paternsRegex);
  if (!paternsMatch) {
    throw new Error("Bloco de dados 'Paterns' não foi encontrado no arquivo.");
  }

  const paternsData = paternsMatch[1];
  const sheetMap = {
    "DS-Right": "DS-R",
    "DS-Left": "DS-L",
    "DS-Center": "DS-C",
    Hazard: "HZD",
    Emergency: "EM1",
    FP1: "FP1",
    FP2: "FP2",
    FP3: "FP3",
    FP4: "FP4",
    AUX1: "AUX1",
    AUX2: "AUX2",
    AUX3: "AUX3",
    AUX4: "AUX4",
  };
  const sheetMapKeys = Object.keys(sheetMap).sort(
    (a, b) => b.length - a.length
  );

  const lines = paternsData.split("\n");
  let currentSheetName = null;
  let firstPopulatedSheet = null;

  // --- NOVA LÓGICA: Ler as cores das colunas diretamente do DOM ---
  const currentColumnColors = {};
  const headerCells = document.querySelectorAll(
    "#macro-header th[data-col-index]"
  );
  const persistentColorRegex = /persistent-color-(\w+)/;

  headerCells.forEach((th) => {
    const colIndex = th.dataset.colIndex;
    const classList = Array.from(th.classList);
    const colorClass = classList.find((cls) => persistentColorRegex.test(cls));

    if (colorClass) {
      const match = colorClass.match(persistentColorRegex);
      if (match && match[1]) {
        currentColumnColors[colIndex] = match[1]; // Ex: { 0: 'red', 2: 'blue' }
      }
    }
  });

  for (const line of lines) {
    const cleanedLine = line.trim();
    if (!cleanedLine) continue;

    let isHeader = false;
    if (cleanedLine.startsWith("//")) {
      for (const key of sheetMapKeys) {
        const headerRegex = new RegExp(`//.*${key.replace("-", "\\-")}`, "i");
        if (headerRegex.test(cleanedLine)) {
          if (/inicio de/i.test(cleanedLine)) {
            currentSheetName = sheetMap[key];
            if (currentSheetName && !populatedSheets.has(currentSheetName)) {
              tempData[currentSheetName] = []; // Inicia um array temporário
              populatedSheets.add(currentSheetName);
            }
            if (!firstPopulatedSheet) firstPopulatedSheet = currentSheetName;
          } else if (/fim de/i.test(cleanedLine)) {
            currentSheetName = null;
          }
          isHeader = true;
          break;
        }
      }
    }

    if (isHeader) continue;

    if (currentSheetName && /^\d|,/.test(cleanedLine)) {
      const dataLine = cleanedLine.replace(/\/\/.*$/, "").trim();
      let parts = dataLine
        .split(",")
        .map((p) => p.trim())
        .filter(Boolean);

      const numericParts = parts.filter((p) =>
        /^(-?\d+|0B[01]+|0X[0-9A-F]+)$/i.test(p)
      );

      if (numericParts.length >= 4) {
        const timeTicks = parseNumericValue(numericParts[0]);
        const flags = /0x80/i.test(parts.join(",")) ? 0x80 : 0x00;
        const byteA = parseNumericValue(numericParts[numericParts.length - 3]);
        const byteB = parseNumericValue(numericParts[numericParts.length - 2]);
        const byteC = parseNumericValue(numericParts[numericParts.length - 1]);

        if (timeTicks !== null) {
          const values = bytesToTableValues(
            byteA,
            byteB,
            byteC,
            currentColumnColors
          );
          tempData[currentSheetName].push({
            ms: (timeTicks * 25).toString(),
            values: values,
          });

          if ((flags & 0x80) !== 0) {
            tempData[currentSheetName].push({
              ms: "0",
              values: Array(24).fill(""),
            });
            currentSheetName = null;
          }
        }
      }
    }
  }

  // Mescla os dados lidos na estrutura fixa
  sheetNames.forEach((name) => {
    const newSheetData = JSON.parse(JSON.stringify(baseSheetStructure));

    if (populatedSheets.has(name)) {
      const parsedRows = tempData[name];
      for (let i = 0; i < parsedRows.length; i++) {
        if (i < newSheetData.length) {
          newSheetData[i] = parsedRows[i];
        }
      }
    }

    newSheetData.forEach((row, index) => {
      row.row = index + 1;
    });

    allMacroData[name] = newSheetData;
  });

  switchSheet(firstPopulatedSheet || "FP1", true);
  renderAllActiveSettings();
  updateConsumptionDisplay();
}

function calculateConsumption(sheetData) {
  const CONSUMO_POR_MODULO = 0.82;
  let consumoTotalPonderado = 0;
  let tempoTotal = 0;
  let picoModulosAtivos = 0;

  if (!sheetData)
    return {
      media: 0,
      pico: 0,
    };

  for (const row of sheetData) {
    const msValue = parseInt(row.ms, 10);
    if (isNaN(msValue) || row.ms.trim() === "") continue;

    if (msValue === 0) break;

    const modulosAtivosNaLinha = row.values.filter(
      (v) => v && v !== "0"
    ).length;
    if (modulosAtivosNaLinha > picoModulosAtivos) {
      picoModulosAtivos = modulosAtivosNaLinha;
    }
    consumoTotalPonderado += modulosAtivosNaLinha * msValue;
    tempoTotal += msValue;
  }

  const mediaModulosAtivos =
    tempoTotal > 0 ? consumoTotalPonderado / tempoTotal : 0;
  const mediaFinalAmperes = mediaModulosAtivos * CONSUMO_POR_MODULO;
  const picoFinalAmperes = picoModulosAtivos * CONSUMO_POR_MODULO;

  return {
    media: mediaFinalAmperes,
    pico: picoFinalAmperes,
  };
}

function updateConsumptionDisplay() {
  allMacroData[activeSheet] = getCurrentTableData();
  const results = calculateConsumption(allMacroData[activeSheet]);
  document.getElementById("output-media-final").textContent = `${results.media
    .toFixed(3)
    .replace(".", ",")} A`;
  document.getElementById("output-pico").textContent = `${results.pico
    .toFixed(3)
    .replace(".", ",")} A`;
}

async function saveSimulationToDatabase() {
  const simulationName = prompt(
    "Por favor, digite um nome para esta simulação:",
    "Simulação Padrão"
  );
  if (!simulationName) {
    alert("Salvamento cancelado.");
    return;
  }
  allMacroData[activeSheet] = getCurrentTableData();
  const consumptionData = calculateConsumption(allMacroData[activeSheet]);

  if (currentLayout) {
    currentLayout.blocks = JSON.parse(JSON.stringify(activeBlockLayout));
  }

  const payload = {
    name: simulationName,
    layout_id: currentLayout.id,
    block_layout: activeBlockLayout,
    column_colors: columnColors,
    data: allMacroData,
    consumoMedia: consumptionData.media,
    consumoPico: consumptionData.pico,
    special_functions: savedSpecialFunctionConfigs,
    jumptabela_settings: savedJumpTabelaSettings,
  };
  try {
    const response = await fetch("save_simulation.php", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });
    const result = await response.json();
    if (result.success) {
      alert("Sucesso! A simulação foi salva no banco de dados!");
    } else {
      alert(
        "O servidor respondeu com um erro: " +
          (result.error || "Erro desconhecido.")
      );
    }
  } catch (error) {
    console.error("ERRO AO SALVAR:", error);
    alert(
      "FALHA! Ocorreu um erro ao tentar salvar a simulação. Verifique o console (F12) para detalhes."
    );
  }
}

async function getGifWorkerURL() {
  if (gifWorkerBlobURL) return gifWorkerBlobURL;
  try {
    const workerScriptURL =
      "https://cdnjs.cloudflare.com/ajax/libs/gif.js/0.2.0/gif.worker.js";
    const response = await fetch(workerScriptURL);
    if (!response.ok)
      throw new Error(`A resposta da rede não foi OK: ${response.statusText}`);
    const workerScriptText = await response.text();
    const blob = new Blob([workerScriptText], {
      type: "application/javascript",
    });
    gifWorkerBlobURL = URL.createObjectURL(blob);
    return gifWorkerBlobURL;
  } catch (error) {
    console.error("Falha ao baixar o script do worker do GIF:", error);
    throw new Error(
      "Não foi possível preparar o componente de geração de GIF. Verifique a conexão com a internet."
    );
  }
}

async function getBackgroundImageDataURL(url) {
  if (backgroundImageDataURL && backgroundImageDataURL.sourceUrl === url) {
    return backgroundImageDataURL.dataUrl;
  }
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "Anonymous";
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = img.width;
      canvas.height = img.height;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0);
      const dataURL = canvas.toDataURL("image/png");
      backgroundImageDataURL = {
        sourceUrl: url,
        dataUrl: dataURL,
      };
      resolve(dataURL);
    };
    img.onerror = () => {
      console.warn(
        `Não foi possível carregar a imagem de fundo '${url}'. O GIF será gerado com fundo sólido.`
      );
      backgroundImageDataURL = {
        sourceUrl: url,
        dataUrl: null,
      };
      resolve(null);
    };
    img.src = url;
  });
}

async function generateGifForSheet(sheetName, progressCallback, workerURL) {
  return new Promise(async (resolve, reject) => {
    // 1. SETUP: Coleta os dados da simulação e prepara o GIF
    const macroData = [];
    for (const row of allMacroData[sheetName]) {
      const ms = parseInt(row.ms, 10);
      if (isNaN(ms) || ms === 0) continue;
      macroData.push(row);
    } // Se não houver dados, cria um quadro padrão para evitar erro
    if (macroData.length === 0) {
      macroData.push({
        ms: "1000",
        values: Array(24).fill(""),
      });
    }

    const gif = new GIF({
      workers: 2,
      quality: 5, // Melhor qualidade
      workerScript: workerURL,
    });

    const simulationArea = document.getElementById("simulation-area");
    const containerRect = simulationArea.getBoundingClientRect(); // Cria um canvas fora da tela com as dimensões exatas da área de simulação

    const canvas = document.createElement("canvas");
    canvas.width = containerRect.width;
    canvas.height = containerRect.height;
    const ctx = canvas.getContext("2d"); // Carrega a imagem de fundo

    const backgroundImage = new Image();
    backgroundImage.src = document.getElementById("siren-model-image").src;
    await new Promise((res) => {
      backgroundImage.onload = res;
      backgroundImage.onerror = res;
    }); // Mapeamento de cores para o canvas

    const colorMap = {
      red: {
        bg: "#ff0000",
        text: "white",
      },
      blue: {
        bg: "#2563eb",
        text: "white",
      },
      white: {
        bg: "#ffffff",
        text: "black",
      },
      yellow: {
        bg: "#ffc000",
        text: "black",
      },
      off: {
        bg: "#d1d5db",
        text: "#1f2937",
      },
    };
    const persistentColorMap = {
      red: colorMap.red,
      blue: colorMap.blue,
      white: colorMap.white,
      yellow: colorMap.yellow,
    };

    try {
      // 2. LOOP DE RENDERIZAÇÃO: Itera sobre cada passo da macro
      for (let i = 0; i < macroData.length; i++) {
        const step = macroData[i];
        const values = step.values.map((v) => parseInt(v) || 0); // Limpa o canvas e desenha o fundo

        ctx.clearRect(0, 0, canvas.width, canvas.height);
        if (backgroundImage.complete && backgroundImage.naturalWidth > 0) {
          const hRatio = canvas.width / backgroundImage.naturalWidth;
          const vRatio = canvas.height / backgroundImage.naturalHeight;
          const ratio = Math.min(hRatio, vRatio) * 0.98; // 98% para bater com o CSS
          const centerShift_x =
            (canvas.width - backgroundImage.naturalWidth * ratio) / 2;
          const centerShift_y =
            (canvas.height - backgroundImage.naturalHeight * ratio) / 2;
          ctx.drawImage(
            backgroundImage,
            0,
            0,
            backgroundImage.naturalWidth,
            backgroundImage.naturalHeight,
            centerShift_x,
            centerShift_y,
            backgroundImage.naturalWidth * ratio,
            backgroundImage.naturalHeight * ratio
          );
        } // 3. DESENHA CADA BLOCO

        activeBlockLayout.forEach((blockData, index) => {
          // Se o bloco não for um placeholder e não estiver ativo, ele não deve ser desenhado.
          const isActive =
            blockData.type !== "placeholder" &&
            (values[parseInt(blockData.text) - 1] || 0) > 0;

          const { text, top, left, width, height, rotate } = blockData;
          const x = (parseFloat(left) / 100) * canvas.width;
          const y = (parseFloat(top) / 100) * canvas.height;

          ctx.save(); // Salva o estado do canvas (sem rotação/translação)
          ctx.translate(x + width / 2, y + height / 2); // Move o ponto de origem para o centro do bloco
          ctx.rotate((rotate * Math.PI) / 180); // Rotaciona
          ctx.translate(-(x + width / 2), -(y + height / 2)); // Move a origem de volta

          if (blockData.type === "placeholder") {
            // Desenha o placeholder cinza
            ctx.fillStyle = "#e0e0e0";
            ctx.strokeStyle = "#bdbdbd";
            ctx.lineWidth = 1;
            ctx.fillRect(x, y, width, height);
            ctx.strokeRect(x, y, width, height);
          } else if (isActive) {
            // Se o bloco de canal estiver ativo, desenha-o
            if (Array.isArray(text)) {
              // Bloco bicolor
              const h = height / 2;
              const val1 = values[parseInt(text[0]) - 1] || 0;
              const color1 =
                persistentColorMap[columnColors[parseInt(text[0]) - 1]] ||
                (val1 ? colorMap.red : colorMap.off);
              ctx.fillStyle = color1.bg;
              ctx.fillRect(x, y, width, h);

              const val2 = values[parseInt(text[1]) - 1] || 0;
              const color2 =
                persistentColorMap[columnColors[parseInt(text[1]) - 1]] ||
                (val2 ? colorMap.red : colorMap.off);
              ctx.fillStyle = color2.bg;
              ctx.fillRect(x, y + h, width, h);

              ctx.fillStyle = color1.text;
              ctx.fillText(text[0], x + width / 2, y + h / 2);
              ctx.fillStyle = color2.text;
              ctx.fillText(text[1], x + width / 2, y + h + h / 2);
            } else {
              // Bloco de cor única
              const persistentColor = columnColors[parseInt(text) - 1];
              const finalColor =
                persistentColorMap[persistentColor] || colorMap.red;

              ctx.fillStyle = finalColor.bg;
              ctx.strokeStyle = "rgba(0,0,0,0.2)";
              ctx.lineWidth = 1;
              ctx.fillRect(x, y, width, height);
              ctx.strokeRect(x, y, width, height);

              ctx.fillStyle = finalColor.text;
              ctx.font = "bold 8px Inter, sans-serif";
              ctx.textAlign = "center";
              ctx.textBaseline = "middle";
              ctx.fillText(text, x + width / 2, y + height / 2);
            }
          }
          ctx.restore();
        }); // 4. ADICIONA O QUADRO AO GIF

        gif.addFrame(canvas, {
          delay: parseInt(step.ms, 10),
          copy: true,
        });
        if (progressCallback) {
          progressCallback(i + 1, macroData.length);
        }
      }
    } catch (error) {
      console.error(
        `Falha ao renderizar quadro para a simulação ${sheetName}:`,
        error
      );
      return reject(
        new Error(`Falha na renderização do quadro para ${sheetName}`)
      );
    } // 5. FINALIZA O GIF

    gif.on("finished", (blob) => {
      resolve(blob);
    });
    gif.render();
  });
}

async function exportAllSimulationsAsGifs() {
  const btn = document.getElementById("export-gif-btn");
  btn.disabled = true;
  btn.innerHTML =
    '<i class="fa fa-spinner fa-spin"></i> Preparando recursos...';

  try {
    const workerURL = await getGifWorkerURL();
    allMacroData[activeSheet] = getCurrentTableData();

    const zip = new JSZip();
    let totalSheetsToProcess = 0;

    // Verifica quais abas têm dados válidos
    for (const sheetName of sheetNames) {
      const data = allMacroData[sheetName];
      if (data && data.length > 0) {
        // Verifica se há alguma linha com ms > 0
        const hasValidData = data.some((row) => {
          const ms = parseInt(row.ms, 10);
          return !isNaN(ms) && ms > 0;
        });
        if (hasValidData) {
          totalSheetsToProcess++;
        }
      }
    }

    if (totalSheetsToProcess === 0) {
      alert("Nenhuma simulação com dados válidos encontrada para exportar.");
      btn.disabled = false;
      btn.innerHTML = '<i class="fa fa-file-image"></i> Exportar GIFs';
      return;
    }

    let processedSheets = 0;

    for (const sheetName of sheetNames) {
      const data = allMacroData[sheetName];
      if (!data || data.length === 0) continue;

      // Verifica se a planilha tem dados válidos
      const hasValidData = data.some((row) => {
        const ms = parseInt(row.ms, 10);
        return !isNaN(ms) && ms > 0;
      });

      if (!hasValidData) {
        console.log(`Pulando aba ${sheetName}: sem dados válidos.`);
        continue; // Pula esta aba
      }

      const progressCallback = (frame, totalFrames) => {
        const percent = Math.round((frame / totalFrames) * 100);
        btn.innerHTML = `<i class="fa fa-spinner fa-spin"></i> Gerando ${sheetName} (${frame}/${totalFrames})...`;
      };

      const gifBlob = await generateGifForSheet(
        sheetName,
        progressCallback,
        workerURL
      );

      if (gifBlob) {
        zip.file(`${sheetName}.gif`, gifBlob);
      }

      processedSheets++;
      // Atualiza o texto do botão após cada aba processada
      if (processedSheets === totalSheetsToProcess) {
        btn.innerHTML = '<i class="fa fa-cog fa-spin"></i> Compactando...';
      }
    }

    // Gera o ZIP com todos os GIFs
    const zipBlob = await zip.generateAsync({
      type: "blob",
    });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(zipBlob);
    link.download = `simulacoes_gifs_${new Date()
      .toLocaleDateString("pt-BR")
      .replace(/\//g, "-")}.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    alert(
      "Pacote de GIFs gerado com sucesso! O download deve começar em breve."
    );
  } catch (error) {
    console.error("ERRO CRÍTICO DURANTE EXPORTAÇÃO DE GIFS:", error);
    alert(
      error.message ||
        "Ocorreu um erro grave ao tentar gerar os GIFs. Verifique o console para detalhes."
    );
  } finally {
    btn.disabled = false;
    btn.innerHTML = '<i class="fa fa-file-image"></i> Exportar GIFs';
    renderLights();
  }
}

function applyColumnColor(colIndex, color) {
  const ALL_COLOR_CLASSES = [
    "persistent-color-red",
    "persistent-color-green",
    "persistent-color-blue",
    "persistent-color-white",
    "persistent-color-yellow",
  ];

  // Seleciona a célula do cabeçalho correspondente
  const headerCell = document.querySelector(
    `#macro-header th[data-col-index="${colIndex}"]`
  );
  const allRows = Array.from(macroTableBody.rows);

  if (color && color !== "clear") {
    columnColors[colIndex] = color;
    const newColorClass = `persistent-color-${color}`;
    allRows.forEach((row) => {
      const cell = row.cells[parseInt(colIndex) + 3];
      // Aplica a cor da coluna no cabeçalho
      if (headerCell) {
        headerCell.classList.remove(...ALL_COLOR_CLASSES);
        headerCell.classList.add(newColorClass);
      }
      if (cell) {
        // Força a atualização da cor da célula para refletir a nova cor da coluna.
        updateCellColor(cell);
      }
    });
  } else {
    delete columnColors[colIndex];
    if (headerCell) {
      // Limpa a cor do cabeçalho
      headerCell.classList.remove(...ALL_COLOR_CLASSES);
    }
    // Limpa a cor da coluna inteira, respeitando as células preenchidas.
    allRows.forEach((row) => {
      const cell = row.cells[parseInt(colIndex) + 3];
      if (cell) {
        cell.classList.remove(...ALL_COLOR_CLASSES);
      }
    });
  }
}

function applyAllColumnColors() {
  for (const colIndex in columnColors) {
    applyColumnColor(colIndex, columnColors[colIndex]);
  }
}

// === INICIALIZAÇÃO ===
// === INICIALIZAÇÃO ===
document.addEventListener("DOMContentLoaded", initializeApp);
Q;
