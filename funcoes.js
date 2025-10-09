// ✅ NOVO: Variável global para armazenar os erros de importação por planilha.
let importErrorsBySheet = {};

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
  // ✅ NOVO: Atribui os elementos do modal de erro
  errorModal = document.getElementById("error-modal");
  importedFilenameDisplay = document.getElementById(
    "imported-filename-display"
  );
  closeErrorModalBtn = document.getElementById("close-error-modal-btn");
  errorListContainer = document.getElementById("error-list-container");
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
      background-color: #00CC00!important; /* Verde Esmeralda */
      color: black;
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
    .block-part-single, .block-part-top, .block-part-bottom {
        transition: background-color 0s;
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
    "Ares",
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
  currentLayoutNameSpan.textContent = formatLayoutName(currentLayout.id);
  document.getElementById("siren-model-image").src = currentLayout.image;
  initializeBlocks(activeBlockLayout);
  stopSimulation();

  // ✅ NOVO: Limpa o nome do arquivo importado ao trocar de layout manualmente
  importedExcelFileName = null;
  importedTxtFileName = null;
  updateImportedFilenamesDisplay();
}

/**
 * ✅ NOVO: Formata o ID de um layout para um nome mais legível e formal.
 * Ex: "primus47p20mRef" se torna "Primus 47" 20 Módulos Refletor".
 * @param {string} layoutId - O ID do layout (ex: "primus47p20mRef").
 * @returns {string} O nome formatado.
 */
function formatLayoutName(layoutId) {
  if (!layoutId) return "Nenhum modelo selecionado";

  // Trata casos especiais primeiro
  if (layoutId.includes("s_slim")) {
    return "S-SLIM Padrão";
  }

  const match = layoutId.match(/^(primus|ares)(\d+p)?(\d+m)(Ref|Col)?/i);

  if (match) {
    const [, produto, tamanho, modulos, tipo] = match;
    const nomeProduto = produto.charAt(0).toUpperCase() + produto.slice(1);
    const nomeTamanho = tamanho ? `${tamanho.replace("p", '"')} ` : "";
    const nomeModulos = `${modulos.replace("m", "")} Módulos`;
    const nomeTipo = tipo
      ? tipo.toLowerCase() === "ref"
        ? "Refletor"
        : "Colimador"
      : "";
    return `${nomeProduto} ${nomeTamanho}${nomeModulos} ${nomeTipo}`
      .replace(/\s+/g, " ")
      .trim();
  }

  // Fallback para IDs que não correspondem ao padrão (ex: nomes antigos)
  return layoutId.replace(/_/g, " ").replace(/\b\w/g, (l) => l.toUpperCase());
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
    numFpSelect.innerHTML += `<option value="${i}" ${i === 4 ? "selected" : ""
      }>${i}</option>`;
    qtdFpDsSelect.innerHTML += `<option value="${i}" ${i === 8 ? "selected" : ""
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
  updateSheetTabsWithWarningIcons();
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
    // ✅ ALTERADO: Adiciona listeners de HOVER para o ícone de aviso.
    button.addEventListener("mouseover", handleSheetTabHover);
    button.addEventListener("mouseout", handleSheetTabMouseOut);
    // Mantém o listener de clique para a funcionalidade principal da aba.
    button.addEventListener("click", (e) =>
      switchSheet(e.target.dataset.sheet)
    );
  });
}

function createTableHeader() {
  macroHeader.innerHTML = "";
  const tr = document.createElement("tr");
  tr.innerHTML += `<th class="column-header">ms/25</th><th class="column-header">ms</th><th class="column-header">#</th>`;
  for (let i = 1; i <= 24; i++) {
    tr.innerHTML += `<th class="column-header" data-col-index="${i - 1
      }" title="Clique com o botão direito para definir uma cor fixa para este canal.">${i}</th>`;
  }
  macroHeader.appendChild(tr);
}

/**
 * ✅ NOVO: Cria uma célula de tempo editável (ms/25 ou ms) com comportamento aprimorado.
 * @param {string} className - A classe CSS para a célula ('ms25-column' ou 'ms-column').
 * @param {string} initialValue - O valor inicial a ser exibido.
 * @param {function} onUpdate - A função de callback a ser chamada quando o valor for alterado.
 * @returns {HTMLTableCellElement} A célula <td> criada.
 */
function createTimeCell(className, initialValue, onUpdate) {
  const cell = document.createElement("td");
  cell.className = className;
  cell.textContent = initialValue;
  cell.contentEditable = true;

  // Ao clicar uma vez, seleciona todo o texto para substituição rápida.
  cell.addEventListener("click", (e) => {
    if (e.detail === 1) {
      // Apenas para clique simples
      document.execCommand("selectAll", false, null);
    }
  });
  // ✅ CORREÇÃO: Ao focar na célula (via clique ou teclado), seleciona todo o texto.
  const selectAllText = () => {
    setTimeout(() => {
      const selection = window.getSelection();
      const range = document.createRange();
      range.selectNodeContents(cell);
      selection.removeAllRanges();
      selection.addRange(range);
    }, 0);
  };

  cell.addEventListener("focus", selectAllText);
  cell.addEventListener("click", selectAllText);

  // Valida a entrada para aceitar apenas números.
  cell.addEventListener("keydown", (e) => {
    // ✅ CORREÇÃO: Previne que a célula fique vazia ao usar Delete/Backspace.
    const selection = window.getSelection();
    if (
      ["Backspace", "Delete"].includes(e.key) &&
      selection.toString() === cell.textContent
    ) {
      e.preventDefault();
      cell.textContent = "0";
      onUpdate();
      return;
    }
    // Permite teclas de controle como Backspace, Delete, setas, Home, End, Ctrl+A/C/V/X/Z/Y
    if (
      [
        "Backspace",
        "Delete",
        "ArrowLeft",
        "ArrowRight",
        "ArrowUp",
        "ArrowDown",
        "Home",
        "End",
        "Tab",
      ].includes(e.key) ||
      (e.ctrlKey &&
        ["a", "c", "v", "x", "z", "y"].includes(e.key.toLowerCase()))
    ) {
      return;
    }
    // Bloqueia qualquer tecla que não seja um número.
    if (!/^\d$/.test(e.key)) {
      e.preventDefault();
    }
  });

  // Sincroniza e formata o valor ao sair da célula.
  cell.addEventListener("blur", () => {
    let value = cell.textContent.trim();
    // Se a célula ficar vazia, define como "0".
    if (value === "") {
      value = "0";
    }
    // Garante que o valor seja um número válido.
    const numericValue = parseInt(value, 10);
    cell.textContent = isNaN(numericValue) ? "0" : numericValue.toString();

    onUpdate(); // Chama a função de sincronização.
    pushStateToUndoHistory(); // Salva no histórico de "desfazer".
  });

  // Chama a sincronização durante a digitação.
  cell.addEventListener("input", onUpdate);

  return cell;
}

function renderTable(data) {
  macroTableBody.innerHTML = "";
  const tableData = data || JSON.parse(JSON.stringify(baseSheetStructure));

  tableData.forEach((rowData, index) => {
    const tr = document.createElement("tr");
    const msValue = parseInt(rowData.ms, 10);
    const ms25Value =
      isNaN(msValue) || rowData.ms === ""
        ? ""
        : msValue === 0
          ? "0"
          : Math.round(msValue / 25);

    // Célula do contador de linha (não editável)
    const counterCell = document.createElement("td");
    counterCell.className = "row-counter-column";
    // ✅ CORREÇÃO: Se rowData.row estiver vazio, usa o índice + 1 como fallback
    counterCell.textContent = rowData.row || index + 1;

    // ✅ REATORAÇÃO: Usa a nova função para criar as células de tempo.
    const msCell = createTimeCell("ms-column", rowData.ms || "0", () =>
      syncTimeValues(ms25Cell, msCell, "from_ms")
    );
    const ms25Cell = createTimeCell("ms25-column", ms25Value, () =>
      syncTimeValues(ms25Cell, msCell, "from_ms25")
    );

    tr.appendChild(ms25Cell);
    tr.appendChild(msCell);
    tr.appendChild(counterCell);

    // Cria as células de dados (1-24)
    const dataValues = rowData.values || [];
    for (let i = 0; i < 24; i++) createEditableCell(tr, dataValues[i] || "", i);
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

/**
 * ✅ CORREÇÃO: Função reintroduzida para criar as células de dados (1-24).
 * Esta função foi removida por engano na refatoração anterior.
 * @param {HTMLTableRowElement} parent - A linha <tr> onde a célula será adicionada.
 * @param {string} value - O valor inicial da célula ('1' ou '').
 * @param {number} colIndex - O índice da coluna (0 a 23).
 */
function createEditableCell(parent, value, colIndex) {
  const td = document.createElement("td");
  td.className = "editable-cell";
  td.contentEditable = true;
  td.dataset.colIndex = colIndex;
  td.textContent = value;
  updateCellColor(td);

  td.addEventListener("keydown", function (e) {
    const colIdx = parseInt(this.dataset.colIndex, 10);
    const channelNum = colIdx + 1;
    // ✅ ALTERADO: Permite a edição com um aviso.
    if (
      e.key === "1" &&
      unconfiguredChannels.has(channelNum) &&
      this.textContent !== "1"
    ) {
      const confirmed = confirm(
        `Atenção: O canal ${channelNum} não está configurado no layout do Excel.\n\nAtivá-lo pode causar um comportamento inesperado.\n\nDeseja continuar?`
      );
      if (!confirmed) {
        e.preventDefault(); // Impede a digitação do "1" se o usuário cancelar.
      } else {
        // Se o usuário confirmar, registra o "erro" para exibir o ícone de aviso.
        const rowNum = parseInt(this.parentElement.cells[2].textContent, 10);
        const error = { sheet: activeSheet, row: rowNum, channel: channelNum };

        if (!importErrorsBySheet[activeSheet]) {
          importErrorsBySheet[activeSheet] = [];
        }
        // Adiciona o erro apenas se ele já não existir para esta célula.
        if (
          !importErrorsBySheet[activeSheet].some(
            (err) => err.row === rowNum && err.channel === channelNum
          )
        ) {
          importErrorsBySheet[activeSheet].push(error);
        }
        updateSheetTabsWithWarningIcons(); // Atualiza a UI imediatamente.
      }
    }
  });

  td.addEventListener("blur", pushStateToUndoHistory);

  td.addEventListener("input", function () {
    const originalText = this.textContent;
    const hasNumber = /[1-9]/.test(originalText);
    const newText = hasNumber ? "1" : "";
    if (originalText !== newText) {
      this.textContent = newText;
    }
    updateCellColor(this);
    updateConsumptionDisplay();

    // ✅ NOVO: Remove o aviso se a célula for limpa.
    const colIdx = parseInt(this.dataset.colIndex, 10);
    const channelNum = colIdx + 1;
    const rowNum = parseInt(this.parentElement.cells[2].textContent, 10);

    if (newText === "" && unconfiguredChannels.has(channelNum)) {
      if (importErrorsBySheet[activeSheet]) {
        // Filtra o array, removendo o erro correspondente a esta célula.
        importErrorsBySheet[activeSheet] = importErrorsBySheet[
          activeSheet
        ].filter((err) => !(err.row === rowNum && err.channel === channelNum));
        // Se não houver mais erros na planilha, remove o ícone.
        updateSheetTabsWithWarningIcons();
      }
    }

    debouncedPushState();
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
  const data = extractedData || fallbackData;
  if (!data) return null;

  let modules = [];
  let productInfo = {};
  let totalModules = 0;

  if (data.mapeamento_cores_modulos || data.pythonData) {
    const pythonData = data.pythonData || data;
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
  } else if (data.layout || data.modules) {
    modules = data.layout || data.modules || [];
    totalModules = modules.length;
    productInfo = data.product_info || { produto: "" };
  }

  if (totalModules === 0) return null;

  const produto = (productInfo.produto || "").toLowerCase();
  const tecnologia = (productInfo.tecnologia || "").toLowerCase();

  console.log(
    `Detectando layout: ${totalModules} módulos ativos, produto: ${produto}, tecnologia: ${tecnologia}`
  );

  // ✅ CORREÇÃO: compara com "ares" em minúsculo
  if (produto.includes("ares")) {
    const isCollimator = tecnologia.includes("eco") || tecnologia.includes("star");
    const layoutSuffix = isCollimator ? "Col" : "Ref";
    const exactLayoutId = `Ares${totalModules}m${layoutSuffix}`;

    const layoutExists = allAvailableLayouts.some((l) => l.id === exactLayoutId);
    if (layoutExists) {
      console.log(`✓ Layout exato detectado: ${exactLayoutId}`);
      return exactLayoutId;
    } else {
      console.error(`❌ Layout inválido! A combinação "${exactLayoutId}" não existe.`);
      return null;
    }
  }

  // Se não for ARES, tenta Primus (opcional, mas mantido para compatibilidade)
  const tamanho = productInfo.tamanho_polegadas || "47";
  const productType = `${tamanho}p`;
  const isCollimator = tecnologia.includes("eco") || tecnologia.includes("star");
  const layoutSuffix = isCollimator ? "Col" : "Ref";
  const exactLayoutId = `primus${productType}${totalModules}m${layoutSuffix}`;

  const layoutExists = allAvailableLayouts.some((l) => l.id === exactLayoutId);
  if (layoutExists) {
    console.log(`✓ Layout exato detectado: ${exactLayoutId}`);
    return exactLayoutId;
  }

  console.error(`❌ Layout inválido! A combinação "${exactLayoutId}" não existe.`);
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
    VD: "verde",
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
          `Range real da planilha: ${worksheet["!ref"] || "A1:Z100"} (${range.e.r + 1
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
                cellValue.includes("Ares") ||
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
                  /SIN VISUAL\s+(PRIMUS|ARES|Ares)\s+(ECO|STAR|PRIME|SALIENT|ULTRA)\s+(\d+)P\s+(\d+)M\s+(\d+)L\s+\(([^)]+)\)/i
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
                    /SIN VISUAL\s+(PRIMUS|ARES|Ares)/i
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
          // CORREÇÃO: Adiciona a lógica para encontrar células mescladas
          const merges = worksheet["!merges"] || [];
          console.log(`Encontradas ${merges.length} células mescladas.`);
          merges.forEach((merge, i) => {
            if (i < 10) {
              // Log das 10 primeiras para não poluir
              const range = XLSX.utils.encode_range(merge);
              const startCell = XLSX.utils.encode_cell(merge.s);
              const value = worksheet[startCell]
                ? worksheet[startCell].v
                : "vazio";
              console.log(`  - Merge ${i + 1}: ${range} | Valor: "${value}"`);
            }
          });
          Object.keys(worksheet).forEach((cellKey) => {
            if (!cellKey.startsWith("!")) {
              const cell = worksheet[cellKey]; // Objeto da célula
              if (cell && cell.v) {
                // ✅ CORREÇÃO: Usar 'cell.v' em vez de 'cell.value'
                const cellValue = String(cell.v);
                if (
                  cellValue.includes("PRIMUS") ||
                  cellValue.includes("ARES") ||
                  cellValue.includes("Ares")
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
        // Mapeamento para a linha PRIMUS (baseado em polegadas)
        const MAPA_LAYOUT_EXCEL_PRIMUS = {
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
          // Mapeamento para 54 Polegadas (a ser preenchido)
          54: {
            posicoes_modulos: [
              "X4:Y5",
              "AC4:AD5",
              "AE4:AF5",
              "AG4:AH5",
              "AI6:AJ7",
              "AK10:AL11",
              "AK14:AL15",
              "AK18:AL19",
              "AI22:AJ23",
              "AG24:AH25",
              "AE24:AF25",
              "AC24:AD25",
              "X24:Y25",
              "Q24:R25",
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
              "Q4:R5",
            ],
            posicoes_canais: {
              "X4:Y5": [2, 0],
              "AC4:AD5": [2, 0],
              "AE4:AF5": [2, 0],
              "AG4:AH5": [2, 0],
              "AI6:AJ7": [2, 0],
              "AK10:AL11": [2, 0],
              "AK14:AL15": [0, -2],
              "AK18:AL19": [-2, 0],
              "AI22:AJ23": [-2, 0],
              "AG24:AH25": [-2, 0],
              "AE24:AF25": [-2, 0],
              "AC24:AD25": [-2, 0],
              "X24:Y25": [-2, 0],
              "Q24:R25": [-2, 0],
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
              "O4:P5": [2, 0],
              "Q4:R5": [2, 0],
            },
          },
          // Mapeamento para 61 Polegadas (a ser preenchido)
          61: {
            posicoes_modulos: [
              "AC4:AD5",
              "AE4:AF5",
              "AG4:AH5",
              "AJ4:AK5",
              "AL4:AM5",
              "AN4:AO5",
              "AP6:AQ7",
              "AR10:AS11",
              "AR14:AS15",
              "AR18:AS19",
              "AP22:AQ23",
              "AN24:AO25",
              "AL24:AM25",
              "AJ24:AK25",
              "AG24:AH25",
              "AE24:AF25",
              "AC24:AD25",
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
              "O4:P5",
              "Q4:R5",
              "S4:T5",
              "V4:W5",
              "X4:Y5",
              "Z4:AA5",
            ],
            posicoes_canais: {
              // Módulos Superiores (Linhas 4-7) - Offset [2, 0]
              "AC4:AD5": [2, 0],
              "AE4:AF5": [2, 0],
              "AG4:AH5": [2, 0],
              "AJ4:AK5": [2, 0],
              "AL4:AM5": [2, 0],
              "AN4:AO5": [2, 0],
              "AP6:AQ7": [2, 0],

              // Laterais Direita - Offsets Específicos
              "AR10:AS11": [2, 0], // Linha 10/11: Canal à direita
              "AR14:AS15": [0, -2], // Linha 14/15: Canal à esquerda
              "AR18:AS19": [-2, 0], // Linha 18/19: Canal acima

              // Módulos Inferiores (Linhas 22-25) - Offset [-2, 0]
              "AP22:AQ23": [-2, 0],
              "AN24:AO25": [-2, 0],
              "AL24:AM25": [-2, 0],
              "AJ24:AK25": [-2, 0],
              "AG24:AH25": [-2, 0],
              "AE24:AF25": [-2, 0],
              "AC24:AD25": [-2, 0],
              "Z24:AA25": [-2, 0],
              "X24:Y25": [-2, 0],
              "V24:W25": [-2, 0],
              "S24:T25": [-2, 0],
              "Q24:R25": [-2, 0],
              "O24:P25": [-2, 0],
              "L24:M25": [-2, 0],
              "J24:K25": [-2, 0],
              "H24:I25": [-2, 0],

              // Laterais Esquerda - Offsets Específicos
              "F22:G23": [-2, 0], // Curva Inferior Esquerda: Canal acima
              "D18:E19": [-2, 0], // Linha 18/19: Canal abaixo
              "D14:E15": [0, 2], // Linha 14/15: Canal à direita
              "D10:E11": [2, 0], // Linha 10/11: Canal à esquerda

              // Módulos Superiores/Curvas (Linhas 6-7, 4-5) - Offset [2, 0]
              "F6:G7": [2, 0],
              "H4:I5": [2, 0],
              "J4:K5": [2, 0],
              "L4:M5": [2, 0],
              "O4:P5": [2, 0],
              "Q4:R5": [2, 0],
              "S4:T5": [2, 0],
              "V4:W5": [2, 0],
              "X4:Y5": [2, 0],
              "Z4:AA5": [2, 0],
            },
          },
        };

        // NOVO: Mapeamento para a linha ARES (baseado em quantidade de módulos)
        // NOVO: Mapeamento ÚNICO para a linha ARES (todas as posições possíveis)
        const MAPA_LAYOUT_EXCEL_ARES = {
          posicoes_modulos: [
            "R4:S5", "T4:U5", "V6:W7", "X8:Y9", "Z8:AA9", "AB10:AC11",
            "AD12:AE13", "AE16:AF17", "AC20:AD21", "Z20:AA21", "X20:Y21", "V18:W19",
            "T16:U17", "R16:S17", "P16:Q17", "N18:O19", "L20:M21", "J20:K21",
            "G20:H21", "E16:F17", "F12:G13", "H10:I11", "J8:K9", "L8:M9", "N6:O7", "P4:Q5"
          ],
          posicoes_canais: {
            "R4:S5": [2, 0], "T4:U5": [2, 0], "V6:W7": [2, 0], "X8:Y9": [2, 0], "Z8:AA9": [2, 0], "AB10:AC11": [2, 0],
            "AD12:AE13": [2, 0], "AE16:AF17": [2, 0], "AC20:AD21": [-2, 0], "Z20:AA21": [-2, 0], "X20:Y21": [-2, 0],
            "V18:W19": [-2, 0], "T16:U17": [-2, 0], "R16:S17": [-2, 0], "P16:Q17": [-2, 0], "N18:O19": [-2, 0],
            "L20:M21": [-2, 0], "J20:K21": [-2, 0], "G20:H21": [-2, 0], "E16:F17": [2, 0], "F12:G13": [2, 0],
            "H10:I11": [2, 0], "J8:K9": [2, 0], "L8:M9": [2, 0], "N6:O7": [2, 0], "P4:Q5": [2, 0]
          }
        };
        // ========================================================================================//
        // =================== FUNÇÃO DE EXTRAÇÃO DO EXCEL (PYTHON BACKEND) ================= //
        // ========================================================================================//

        // --- LÓGICA DE SELEÇÃO DE MAPEAMENTO (PRIMUS vs ARES) ---
        let layoutConfig = null;
        const produtoDetectado = (productInfo.produto || "").toLowerCase();

        if (produtoDetectado.includes("ares")) { // ← minúsculo!
          layoutConfig = MAPA_LAYOUT_EXCEL_ARES; // usa o mapa único
          console.log("Produto ARES detectado. Usando mapa único de posições.");
        } else {
          const tamanhoFinal = tamanhoDetectadoPelaPlanilha || productInfo.tamanho_polegadas || "47";
          layoutConfig = MAPA_LAYOUT_EXCEL_PRIMUS[tamanhoFinal];
          console.log(`Produto PRIMUS/Padrão detectado. Usando mapa para ${tamanhoFinal} polegadas.`);
        }

        if (!layoutConfig || layoutConfig.posicoes_modulos.length === 0) {
          const erroMsg = `Configuração de layout para o produto detectado não foi encontrada ou está vazia. Produto: ${produtoDetectado}, Módulos/Tamanho: ${produtoDetectado.includes("ares")
            ? productInfo.quantidade_modulos
            : productInfo.tamanho_polegadas
            }`;
          console.error(erroMsg);
          // Opcional: pode rejeitar a promessa aqui se for um erro crítico
          // return reject(new Error(erroMsg));
        }

        const posicoes_modulos =
          layoutConfig?.posicoes_modulos ||
          MAPA_LAYOUT_EXCEL_PRIMUS["47"].posicoes_modulos;
        const POSICOES_CANAIS =
          layoutConfig?.posicoes_canais ||
          MAPA_LAYOUT_EXCEL_PRIMUS["47"].posicoes_canais;

        const mapeamento_cores_modulos = [];
        console.log("Iniciando processamento de módulos...");

        // Estrutura para rastrear células já processadas como parte de um par dual color
        const processedCells = new Set();

        // Função para encontrar se uma célula faz parte de uma mesclagem
        const findMerge = (r, c) => {
          const merges = worksheet["!merges"] || [];
          for (const merge of merges) {
            if (
              r >= merge.s.r &&
              r <= merge.e.r &&
              c >= merge.s.c &&
              c <= merge.e.c
            ) {
              return merge;
            }
          }
          return null;
        };
        // Processa cada módulo
        posicoes_modulos.forEach((range_str, index) => {
          const modulo_num = index + 1;
          const mapeamento = [];

          try {
            // === NOVA LÓGICA SIMPLIFICADA DE CLASSIFICAÇÃO DE MÓDULOS ===
            const range = XLSX.utils.decode_range(range_str);

            // 1. Escaneia toda a área 2x2 do módulo e coleta todas as cores únicas
            const coresUnicas = new Set(); // Para detectar cores únicas
            const coresDetalhadas = []; // Para armazenar posições detalhadas

            console.log(`\n=== Módulo ${modulo_num} (${range_str}) ===`);
            console.log(
              `Escaneando área 2x2: linhas ${range.s.r}-${range.e.r}, colunas ${range.s.c}-${range.e.c}`
            );

            // Varre toda a área 2x2 do módulo
            for (let row = range.s.r; row <= range.e.r; row++) {
              for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];

                if (cell && cell.v) {
                  const textoOriginal = String(cell.v).trim();
                  const texto_celula = textoOriginal.toUpperCase();

                  console.log(`  Célula ${cellAddress}: "${textoOriginal}"`);

                  // Verifica se contém cores válidas (VM, AZ, BR, AB)
                  if (COLOR_MAP[texto_celula]) {
                    // Cor única válida
                    coresUnicas.add(texto_celula);
                    coresDetalhadas.push({
                      codigo: texto_celula,
                      nome: COLOR_MAP[texto_celula],
                      posicao: { r: row, c: col },
                      celula: cellAddress,
                    });
                    console.log(
                      `    ✓ Cor válida: ${texto_celula} (${COLOR_MAP[texto_celula]})`
                    );
                  }
                  // Verifica se contém múltiplas cores separadas por delimitadores
                  else if (
                    texto_celula.includes("+") ||
                    texto_celula.includes("/") ||
                    texto_celula.includes("|")
                  ) {
                    const coresSeparadas = texto_celula
                      .split(/[+\/|]/)
                      .map((c) => c.trim());
                    console.log(
                      `    Cores separadas encontradas: ${coresSeparadas.join(
                        ", "
                      )}`
                    );

                    coresSeparadas.forEach((cor, index) => {
                      if (COLOR_MAP[cor]) {
                        coresUnicas.add(cor);
                        coresDetalhadas.push({
                          codigo: cor,
                          nome: COLOR_MAP[cor],
                          posicao: { r: row, c: col },
                          celula: cellAddress,
                          indice: index,
                        });
                        console.log(
                          `      ✓ Cor ${index + 1}: ${cor} (${COLOR_MAP[cor]})`
                        );
                      }
                    });
                  }
                }
              }
            }

            const numCores = coresUnicas.size;
            console.log(`\n📊 Classificação do módulo ${modulo_num}:`);
            console.log(`   - Número de cores únicas: ${numCores}`);
            console.log(
              `   - Cores encontradas: [${Array.from(coresUnicas).join(", ")}]`
            );
            console.log(`   - Tipo: ${numCores}-color`);

            // 2. Se encontrou cores, processa o mapeamento baseado na classificação
            if (numCores > 0) {
              const canal_info = POSICOES_CANAIS[range_str];

              if (!canal_info) {
                console.log(
                  `   ⚠ Configuração de canal não encontrada para ${range_str}`
                );
                return;
              }

              const [row_offset, col_offset] = canal_info;
              console.log(
                `   - Offset do canal: [${row_offset}, ${col_offset}]`
              );

              if (numCores === 1) {
                // === MÓDULO 1-COLOR ===
                const corInfo = coresDetalhadas[0];

                // Para 1-color, usa a primeira célula do módulo + offset
                const nova_row = range.s.r + row_offset;
                const nova_col = range.s.c + col_offset;

                const coord_canal = XLSX.utils.encode_cell({
                  r: nova_row,
                  c: nova_col,
                });
                const canal_cell = worksheet[coord_canal];
                const numero_canal =
                  canal_cell && canal_cell.v && !isNaN(canal_cell.v)
                    ? parseInt(canal_cell.v, 10)
                    : null;

                console.log(`   🔵 1-Color: ${corInfo.nome}`);
                console.log(
                  `      Módulo range: ${range_str} (${range.s.r}, ${range.s.c})`
                );
                console.log(`      Offset: [${row_offset}, ${col_offset}]`);
                console.log(
                  `      Posição do canal: linha ${nova_row}, coluna ${nova_col}`
                );
                console.log(
                  `      Canal em ${coord_canal}: ${numero_canal || "não encontrado"
                  }`
                );

                mapeamento.push({
                  cor: corInfo.nome,
                  canal: numero_canal,
                  canal_encontrado_em: coord_canal,
                });
              } else {
                // === MÓDULOS MULTICOLOR (2, 3, 4 cores) ===
                console.log(
                  `   🌈 ${numCores}-Color: processamento sequencial`
                );

                // Ordena as cores por posição: linha primeiro, depois coluna
                coresDetalhadas.sort((a, b) => {
                  if (a.posicao.r !== b.posicao.r) {
                    return a.posicao.r - b.posicao.r; // Linha: cima para baixo
                  }
                  return a.posicao.c - b.posicao.c; // Coluna: esquerda para direita
                });

                // Remove duplicatas mantendo a primeira ocorrência
                const coresOrdenadas = [];
                const jaProcesadas = new Set();

                coresDetalhadas.forEach((corInfo) => {
                  if (!jaProcesadas.has(corInfo.codigo)) {
                    jaProcesadas.add(corInfo.codigo);
                    coresOrdenadas.push(corInfo);
                  }
                });

                console.log(`   📍 Ordem de processamento:`);
                coresOrdenadas.forEach((cor, index) => {
                  console.log(
                    `      ${index + 1}. ${cor.codigo} (${cor.nome}) em ${cor.celula
                    }`
                  );
                });

                console.log(`   📍 Cálculo dos canais:`);
                console.log(
                  `      Módulo range: ${range_str} (${range.s.r}, ${range.s.c})`
                );
                console.log(`      Offset: [${row_offset}, ${col_offset}]`);

                // Calcula canais com lógica específica para bicolor
                coresOrdenadas.forEach((corInfo, index) => {
                  let nova_row, nova_col;

                  if (numCores === 2) {
                    // === LÓGICA ESPECIAL PARA BICOLOR ===
                    if (index === 0) {
                      // Primeira cor: usa sua própria posição + offset
                      nova_row = corInfo.posicao.r + row_offset;
                      nova_col = corInfo.posicao.c + col_offset;
                    } else {
                      // Segunda cor: linha da primeira cor + coluna da primeira cor + 1 + offset
                      const primeiraCor = coresOrdenadas[0];
                      nova_row = primeiraCor.posicao.r + row_offset;
                      nova_col = primeiraCor.posicao.c + 1 + col_offset;
                      console.log(
                        `      BICOLOR: Segunda cor usa linha da primeira (${primeiraCor.posicao.r
                        }) + coluna da primeira + 1 (${primeiraCor.posicao.c + 1
                        })`
                      );
                    }
                  } else {
                    // === LÓGICA PADRÃO PARA 1, 3, 4 CORES ===
                    // Aplica o offset diretamente à posição real da cor
                    nova_row = corInfo.posicao.r + row_offset;
                    nova_col = corInfo.posicao.c + col_offset;
                  }

                  const coord_canal = XLSX.utils.encode_cell({
                    r: nova_row,
                    c: nova_col,
                  });
                  const canal_cell = worksheet[coord_canal];
                  const numero_canal =
                    canal_cell && canal_cell.v && !isNaN(canal_cell.v)
                      ? parseInt(canal_cell.v, 10)
                      : null;

                  console.log(
                    `      Cor ${index + 1} (${corInfo.nome}) em ${corInfo.celula
                    } (${corInfo.posicao.r}, ${corInfo.posicao.c
                    }): Canal em ${coord_canal} (${nova_row}, ${nova_col}) = ${numero_canal || "não encontrado"
                    }`
                  );

                  mapeamento.push({
                    cor: corInfo.nome,
                    canal: numero_canal,
                    canal_encontrado_em: coord_canal,
                  });
                });
              }
            } else {
              console.log(
                `   ❌ Nenhuma cor válida encontrada no módulo ${modulo_num}`
              );
            }
            // === FIM DA NOVA LÓGICA SIMPLIFICADA ===
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

// ✅ NOVA FUNÇÃO: Mostra o popup de cores baseado na posição da célula de referência
function showColorPopup(referenceCell) {
  const colorPopup = document.getElementById("color-palette-popup");

  if (colorPopup && referenceCell) {
    const headerRect = referenceCell.getBoundingClientRect();

    colorPopup.style.display = "block";
    // Posiciona o popup abaixo do cabeçalho de referência
    colorPopup.style.top = `${headerRect.bottom + window.scrollY + 5}px`;
    colorPopup.style.left = `${headerRect.left + window.scrollX}px`;

    // ✅ NOVO: Força alguns estilos para garantir visibilidade
    colorPopup.style.visibility = "visible";
    colorPopup.style.opacity = "1";
    colorPopup.style.zIndex = "10000"; // Aumentando para garantir que fique por cima
    colorPopup.style.position = "absolute";

    // ✅ NOVO: Força uma atualização visual do elemento
    colorPopup.offsetHeight; // Força um reflow do elemento

    // ✅ NOVO: Marca o timestamp de quando o popup foi mostrado para evitar fechamento imediato
    colorPopup.dataset.shownAt = Date.now();
  }
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

  // ✅ NOVO: Listener para fechar o modal de erro
  if (closeErrorModalBtn) {
    closeErrorModalBtn.addEventListener("click", () => {
      errorModal.style.display = "none";
    });
  }

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

      // ✅ ALTERADO: Armazena o nome do arquivo Excel e atualiza o display.
      importedExcelFileName = file.name;
      updateImportedFilenamesDisplay();

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
              applyColumnColorOptimized(i.toString(), "clear");
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
                      applyColumnColorOptimized(
                        colIndex.toString(),
                        translatedColor
                      );
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
            // ✅ NOVO: Identifica os canais que NÃO foram usados no Excel.
            const todosCanaisPossiveis = new Set(
              Array.from({ length: 24 }, (_, i) => i + 1)
            );
            const canaisUsados = new Set();
            extractedResult.pythonData.mapeamento_cores_modulos.forEach(
              (mod) => {
                mod.mapeamento.forEach((map) => {
                  if (map.canal !== null && map.canal > 0) {
                    canaisUsados.add(map.canal);
                  }
                });
              }
            );
            const canaisNaoUsados = [...todosCanaisPossiveis].filter(
              (c) => !canaisUsados.has(c)
            );
            console.log(
              "Canais usados no Excel:",
              [...canaisUsados].sort((a, b) => a - b)
            );
            console.log(
              "Canais NÃO USADOS que serão exibidos como extras:",
              canaisNaoUsados.sort((a, b) => a - b)
            );
            // ✅ NOVO: Armazena os canais não configurados em uma variável global para uso posterior.
            unconfiguredChannels = new Set(canaisNaoUsados);

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
            currentLayoutNameSpan.textContent = formatLayoutName(
              currentLayout.id
            );
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
        // ✅ CORREÇÃO: Primeiro para a simulação para limpar o estado anterior.
        stopSimulation();
        // ✅ CORREÇÃO: Garante que todos os blocos fiquem visíveis no modo de edição.
        document
          .querySelectorAll("#blocks-container .block-container")
          .forEach((block) => (block.style.opacity = "1"));
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
        // ✅ NOVO: Ao sair do modo de edição, restaura a visualização da linha selecionada.
        const selectedRow = macroTableBody.querySelector("tr.selected-row");
        if (selectedRow) {
          previewRow(selectedRow.rowIndex - 1);
        }
      }
    });
  }

  // ✅ NOVO: Adiciona a lógica de seleção de coluna com clique esquerdo.
  // Lógica aprimorada para incluir clique, arrasto, ctrl e shift.
  if (macroHeader) {
    let isDraggingSelection = false;
    let startDragIndex = -1;

    macroHeader.addEventListener("mousedown", (e) => {
      const startCell = e.target.closest("th.column-header[data-col-index]");

      if (!startCell) {
        return;
      }

      e.preventDefault(); // Previne a seleção de texto padrão durante o arrasto.
      isDraggingSelection = true;
      startDragIndex = parseInt(startCell.dataset.colIndex, 10);

      // Lógica de clique inicial (mousedown)
      if (e.shiftKey && lastSelectedColumn !== -1) {
        const start = Math.min(lastSelectedColumn, startDragIndex);
        const end = Math.max(lastSelectedColumn, startDragIndex);
        if (!e.ctrlKey) selectedColumns.clear();
        for (let i = start; i <= end; i++) selectedColumns.add(i);
      } else if (e.ctrlKey) {
        if (selectedColumns.has(startDragIndex)) {
          selectedColumns.delete(startDragIndex);
        } else {
          selectedColumns.add(startDragIndex);
        }
      } else {
        selectedColumns.clear();
        selectedColumns.add(startDragIndex);
      }

      lastSelectedColumn = startDragIndex;
      updateSelectedColumnsVisuals();

      // Handler para o movimento do mouse
      const onMouseMove = (moveEvent) => {
        if (!isDraggingSelection) return;

        const hoverCell = moveEvent.target.closest(
          "th.column-header[data-col-index]"
        );

        if (hoverCell) {
          const hoverIndex = parseInt(hoverCell.dataset.colIndex, 10);
          const start = Math.min(startDragIndex, hoverIndex);
          const end = Math.max(startDragIndex, hoverIndex);

          // ✅ OTIMIZAÇÃO: Apenas atualiza o visual dos cabeçalhos durante o arrasto.
          document
            .querySelectorAll("#macro-header th[data-col-index]")
            .forEach((th) => {
              const idx = parseInt(th.dataset.colIndex, 10);
              th.classList.toggle(
                "column-selected",
                idx >= start && idx <= end
              );
            });
        }
      };

      // Handler para quando o botão do mouse é solto
      const onMouseUp = (upEvent) => {
        isDraggingSelection = false;
        const endCell = upEvent.target.closest(
          "th.column-header[data-col-index]"
        );

        if (endCell) {
          lastSelectedColumn = parseInt(endCell.dataset.colIndex, 10);
          // ✅ CORREÇÃO: Adiciona as colunas do intervalo final à seleção.

          const start = Math.min(startDragIndex, lastSelectedColumn);
          const end = Math.max(startDragIndex, lastSelectedColumn);

          // Limpa a seleção anterior se não estiver usando Ctrl/Shift para adicionar.
          if (!upEvent.ctrlKey && !upEvent.shiftKey) {
            selectedColumns.clear();
          }
          // Adiciona todas as colunas no intervalo ao conjunto de seleção.
          for (let i = start; i <= end; i++) selectedColumns.add(i);
        }

        // ✅ OTIMIZAÇÃO: Atualiza a seleção final (cabeçalho e corpo) apenas uma vez.
        updateSelectedColumnsVisuals();

        // ✅ CORREÇÃO: Mostra o popup de cores após a seleção ser finalizada e visualizada.
        // A lógica foi movida para depois de 'updateSelectedColumnsVisuals' para garantir que funcione sempre.
        // Usa a última célula selecionada (seja pelo clique ou pelo fim do arrasto) como referência.
        const referenceCell = endCell || startCell;

        if (selectedColumns.size > 0) {
          showColorPopup(referenceCell);
        }

        // Remove os listeners para não sobrecarregar o navegador
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      };

      // Adiciona os listeners temporários
      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    });

    // ✅ NOVO: Adiciona listener de clique para garantir que o popup apareça em cliques simples
    macroHeader.addEventListener("click", (e) => {
      const clickedCell = e.target.closest("th.column-header[data-col-index]");
      if (!clickedCell) return;

      // Se há colunas selecionadas, mostra o popup
      if (selectedColumns.size > 0) {
        showColorPopup(clickedCell);
      }
    });
  }

  if (macroTableBody) {
    macroTableBody.addEventListener("mousedown", (e) => {
      // Desseleciona colunas ao clicar em uma célula, mas permite que a seleção de texto na célula funcione.
      if (!e.target.closest("#macro-header")) {
        // Atraso mínimo para garantir que o clique não seja parte de uma ação de seleção de texto
        setTimeout(() => {
          selectedColumns.clear();
          updateSelectedColumnsVisuals();
        }, 50);
      }
    });
  }

  /**
   * ✅ NOVO: Função para atualizar o visual de todas as colunas selecionadas.
   */
  function updateSelectedColumnsVisuals() {
    // Primeiro, limpa todas as seleções existentes na tabela
    document
      .querySelectorAll(".column-selected")
      .forEach((el) => el.classList.remove("column-selected"));

    const tableBody = document.getElementById("macro-body");
    const headerCells = document.querySelectorAll(
      "#macro-header th[data-col-index]"
    );

    // Itera sobre as colunas que devem estar selecionadas
    selectedColumns.forEach((colIdx) => {
      // Destaca o cabeçalho
      const header = Array.from(headerCells).find(
        (th) => parseInt(th.dataset.colIndex, 10) === colIdx
      );
      if (header) header.classList.add("column-selected");

      // Destaca as células do corpo
      if (tableBody) {
        const rows = tableBody.getElementsByTagName("tr");
        for (let i = 0; i < rows.length; i++) {
          const cells = rows[i].getElementsByTagName("td");
          for (let j = 0; j < cells.length; j++) {
            if (parseInt(cells[j].dataset.colIndex, 10) === colIdx) {
              cells[j].classList.add("column-selected");
              break;
            }
          }
        }
      }
    });
  }

  // ✅ NOVO: Listener para a paleta de cores que substitui o menu de contexto.
  const colorPalette = document.getElementById("color-palette");
  if (colorPalette) {
    colorPalette.addEventListener("click", (e) => {
      const swatch = e.target.closest(".color-swatch");
      if (!swatch) return;

      const color = swatch.dataset.color;

      if (selectedColumns.size > 0) {
        // Aplica a cor a todas as colunas selecionadas.
        selectedColumns.forEach((colIdx) => {
          applyColumnColorOptimized(colIdx, color); // ✅ OTIMIZAÇÃO: Usa a nova função otimizada.
        });
        // ✅ NOVO: Esconde o popup após aplicar a cor.
        const colorPopup = document.getElementById("color-palette-popup");
        if (colorPopup) {
          colorPopup.style.display = "none";
        }
        // Limpa a seleção após a aplicação da cor.
        selectedColumns.clear();
        updateSelectedColumnsVisuals();
      } else {
        alert(
          "Selecione uma ou mais colunas clicando nos cabeçalhos (1-24) antes de aplicar uma cor."
        );
      }
    });

    // ✅ NOVO: Esconde o popup se clicar fora dele.
    document.addEventListener("click", (e) => {
      const colorPopup = document.getElementById("color-palette-popup");
      if (
        colorPopup &&
        !colorPopup.contains(e.target) &&
        !e.target.closest("th.column-header") &&
        !e.target.closest("#color-palette-popup")
      ) {
        // ✅ NOVO: Proteção temporal - não esconde o popup se foi mostrado há menos de 100ms
        const shownAt = parseInt(colorPopup.dataset.shownAt) || 0;
        const timeSinceShown = Date.now() - shownAt;

        if (timeSinceShown < 100) {
          return;
        }

        colorPopup.style.display = "none";
      }
    });
  }

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
    activeBlock.style.left = `${((action.initialLeft + dx) / containerRect.width) * 100
      }%`;
    activeBlock.style.top = `${((action.initialTop + dy) / containerRect.height) * 100
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
  executeMacroBtn.textContent = "PARAR SIMULAÇÃO";
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

    const values = step.values.map((v) => parseInt(v, 10) || 0); // ✅ NOVO: Passa o número da linha atual para renderLights.
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
    executeMacroBtn.textContent = "SIMULAR";
    executeMacroBtn.style.backgroundColor = "#f39c12";
  }
  clearTimeout(simulationStopper);
  simulationStopper = null;
  renderLights();
  document
    .querySelectorAll("#macro-body tr.selected-row")
    .forEach((tr) => tr.classList.remove("selected-row"));
}

function renderLights(macroValues = [], currentRowNumber = -1) {
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

    if (finalColorValue > 0) {
      if (blockContainer) {
        // Torna o bloco visível apenas se ele estiver ativo neste passo.
        // ✅ CORREÇÃO: Não altera a opacidade se estiver em modo de edição.
        if (!isEditMode) {
          // A opacidade final será o brilho definido (ex: 1.0, 0.75, 0.5).
          blockContainer.style.opacity = brightness;
        }
      }
      // Se o bloco não estiver ativo, ele ficará invisível
      else if (blockContainer) {
        if (!isEditMode) {
          blockContainer.style.opacity = "0";
        }
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
    } else {
      // Se o valor final da cor for 0, o bloco deve ficar invisível.
      // ✅ CORREÇÃO: Não altera a opacidade se estiver em modo de edição.
      if (blockContainer && !isEditMode) {
        blockContainer.style.opacity = "0";
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
    renderLights(values, rowData.row); // ✅ NOVO: Passa o número da linha para a pré-visualização.
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
  const result = Array.from(macroTableBody.rows).map((row, index) => {
    const cells = Array.from(row.cells);
    const rowData = {
      row: cells[2].textContent.trim(),
      ms: cells[1].textContent.trim(),
      values: cells.slice(3).map((cell) => cell.textContent.trim()),
    };

    return rowData;
  });

  return result;
}

function switchSheet(newSheetName, preventSave = false) {
  if (activeSheet === newSheetName) {
    return;
  }

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
    // ✅ ALTERADO: Listener para deselecionar a COLUNA (e não a linha) ao clicar fora da tabela.
    document.addEventListener("click", (e) => {
      // Verifica se o clique foi fora da tabela e dos seus controles principais.
      if (
        !macroTableBody.contains(e.target) &&
        !e.target.closest(".sheet-selector-btn") &&
        !e.target.closest("#macro-header") &&
        !e.target.closest("#color-palette-popup")
      ) {
        // Apenas limpa a seleção de colunas, mantendo a linha selecionada.
        if (selectedColumns.size > 0) {
          selectedColumns.clear();
          updateSelectedColumnsVisuals();
        }
        // A linha selecionada (tr.selected-row) não é mais removida aqui.
        // A pré-visualização (renderLights()) também não é limpa.
      }
    });
  if (macroTableBody)
    macroTableBody.addEventListener("click", (e) => {
      const targetCell = e.target.closest("td");
      if (!targetCell) return;

      if (isSimulationRunning) return;

      // ✅ NOVO: Se a célula for editável, foca nela
      if (targetCell.hasAttribute("contenteditable")) {
        // Força o foco imediatamente
        targetCell.focus();

        // ✅ NOVO: Força o foco novamente após um pequeno delay para garantir que não seja perdido
        setTimeout(() => {
          targetCell.focus();
        }, 10);

        // ✅ NOVO: Posiciona o cursor no final da célula
        const range = document.createRange();
        const sel = window.getSelection();
        range.selectNodeContents(targetCell);
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
      }

      // ✅ CORREÇÃO: Calcula o índice da linha clicada e chama a função de destaque.
      const rowElement = targetCell.parentElement;
      const rowIndex = Array.from(rowElement.parentElement.children).indexOf(
        rowElement
      );
      if (rowIndex !== -1) {
        highlightCurrentRow(rowIndex);
      }
    });

  if (macroTableBody) {
    // ✅ OTIMIZAÇÃO: Lógica de seleção de área (arrastar) refeita para melhor performance.
    let selectionBox = null; // Elemento visual da caixa de seleção

    macroTableBody.addEventListener("mousedown", (e) => {
      const targetCell = e.target.closest('td[contenteditable="true"]');
      if (!targetCell) return;

      // Previne a seleção de texto padrão do navegador durante o arrasto
      e.preventDefault();

      isMouseDown = true;
      startCell = targetCell;
      document.body.style.userSelect = "none";

      // Limpa a seleção anterior
      clearSelection();

      // Cria a caixa de seleção visual
      if (!selectionBox) {
        selectionBox = document.createElement("div");
        selectionBox.style.position = "absolute";
        selectionBox.style.border = "1px dashed #007bff";
        selectionBox.style.backgroundColor = "rgba(0, 123, 255, 0.2)";
        selectionBox.style.pointerEvents = "none"; // Ignora eventos do mouse
        document.body.appendChild(selectionBox);
      }

      const startRect = startCell.getBoundingClientRect();
      const startX = startRect.left + window.scrollX;
      const startY = startRect.top + window.scrollY;

      selectionBox.style.left = `${startX}px`;
      selectionBox.style.top = `${startY}px`;
      selectionBox.style.width = "0px";
      selectionBox.style.height = "0px";
      selectionBox.style.display = "block";

      // ✅ NOVO: Pega o container da grade para limitar a seleção.
      const gridContainer = document.querySelector(".excel-grid");
      const gridRect = gridContainer.getBoundingClientRect();

      // Handler para o movimento do mouse (apenas atualiza a caixa visual)
      const onMouseMove = (moveEvent) => {
        if (!isMouseDown || !startCell) return;

        // ✅ CORREÇÃO: Limita as coordenadas do mouse às bordas do container da tabela.
        let currentX = moveEvent.clientX;
        let currentY = moveEvent.clientY;

        currentX = Math.max(gridRect.left, Math.min(currentX, gridRect.right));
        currentY = Math.max(gridRect.top, Math.min(currentY, gridRect.bottom));

        currentX += window.scrollX;
        currentY += window.scrollY;

        const left = Math.min(startX, currentX);
        const top = Math.min(startY, currentY);
        const width = Math.abs(startX - currentX);
        const height = Math.abs(startY - currentY);

        selectionBox.style.left = `${left}px`;
        selectionBox.style.top = `${top}px`;
        selectionBox.style.width = `${width}px`;
        selectionBox.style.height = `${height}px`;
      };

      // Handler para quando o botão do mouse é solto (calcula a seleção final)
      const onMouseUp = (upEvent) => {
        isMouseDown = false;
        document.body.style.userSelect = "";
        if (selectionBox) selectionBox.style.display = "none";

        const endCell = upEvent.target.closest('td[contenteditable="true"]');
        if (!endCell) {
          // Se soltar fora da tabela, seleciona apenas a célula inicial
          selection.add(startCell);
          startCell.classList.add("cell-selected");
        } else {
          // Calcula a área final e seleciona as células
          const startRowIndex = startCell.parentElement.rowIndex;
          const endRowIndex = endCell.parentElement.rowIndex;
          const startCellIndex = startCell.cellIndex;
          const endCellIndex = endCell.cellIndex;

          const minRow = Math.min(startRowIndex, endRowIndex);
          const maxRow = Math.max(startRowIndex, endRowIndex);
          const minCol = Math.min(startCellIndex, endCellIndex);
          const maxCol = Math.max(startCellIndex, endCellIndex);

          for (let i = minRow; i <= maxRow; i++) {
            const row = macroTableBody.rows[i - 1]; // Ajuste de índice
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
        }

        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      };

      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    });
  }
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

    // ✅ ALTERADO: Deseleciona a COLUNA (e não a linha) ao pressionar a tecla 'Escape'.
    if (e.key === "Escape") {
      e.preventDefault(); // Previne comportamentos padrão do navegador.

      // Apenas limpa a seleção de colunas.
      if (selectedColumns.size > 0) {
        selectedColumns.clear();
        updateSelectedColumnsVisuals();
      }
      // A linha selecionada (tr.selected-row) não é mais removida.
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

    // ✅ PRIORIDADE: Intercepta navegação por setas ANTES de qualquer outra verificação
    if (
      ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Enter"].includes(
        e.key
      )
    ) {
      const isMacroTableCell =
        activeEl &&
        activeEl.tagName === "TD" &&
        activeEl.hasAttribute("contenteditable") &&
        activeEl.closest("#macro-body");

      if (isMacroTableCell) {
        // ✅ IMPORTANTE: Para a propagação imediatamente para evitar interferência
        e.stopPropagation();
        e.preventDefault();

        const currentCell = activeEl;
        const currentRow = currentCell.parentElement;
        const currentCellIndex = currentCell.cellIndex;
        let nextCell;

        switch (e.key) {
          case "Enter":
          case "ArrowDown":
            if (currentRow.nextElementSibling) {
              nextCell = currentRow.nextElementSibling.cells[currentCellIndex];
            }
            break;
          case "ArrowUp":
            if (currentRow.previousElementSibling) {
              nextCell =
                currentRow.previousElementSibling.cells[currentCellIndex];
            }
            break;
          case "ArrowLeft":
            // Só navega se o cursor estiver no início da célula
            if (window.getSelection().anchorOffset === 0) {
              nextCell = currentCell.previousElementSibling;
              // Pula colunas não editáveis (ms/25, ms, #)
              while (nextCell && !nextCell.hasAttribute("contenteditable")) {
                nextCell = nextCell.previousElementSibling;
              }
            } else {
              return; // Permite edição normal na célula
            }
            break;
          case "ArrowRight":
            // Só navega se o cursor estiver no final da célula ou célula vazia
            if (
              window.getSelection().anchorOffset ===
              currentCell.textContent.length ||
              currentCell.textContent.length === 0
            ) {
              nextCell = currentCell.nextElementSibling;
              // Pula colunas não editáveis
              while (nextCell && !nextCell.hasAttribute("contenteditable")) {
                nextCell = nextCell.nextElementSibling;
              }
            } else {
              return; // Permite edição normal na célula
            }
            break;
        }

        if (nextCell && nextCell.hasAttribute("contenteditable")) {
          nextCell.focus();

          // ✅ NOVO: Posiciona o cursor no início da célula
          setTimeout(() => {
            const range = document.createRange();
            const sel = window.getSelection();
            range.setStart(nextCell, 0);
            range.collapse(true);
            sel.removeAllRanges();
            sel.addRange(range);
          }, 1);

          clearSelection();
          selection.add(nextCell);
          nextCell.classList.add("cell-selected");

          // ✅ NOVO: Atualiza a linha selecionada e a pré-visualização ao navegar com as setas
          const nextRow = nextCell.parentElement;
          const rowIndex = Array.from(nextRow.parentElement.children).indexOf(
            nextRow
          );
          highlightCurrentRow(rowIndex);
        }
        return; // Impede que o evento continue para outros elementos
      }
    }

    // ✅ VERIFICAÇÃO SECUNDÁRIA: Para outras tabelas (function-table)

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
      // ✅ CORREÇÃO: Ação de apagar o conteúdo das células selecionadas.
      if (e.key === "Delete" || e.key === "Backspace") {
        e.preventDefault();
        pushStateToUndoHistory();
        selection.forEach((cell) => {
          if (cell.isContentEditable) {
            cell.textContent = "";
            updateCellColor(cell); // Atualiza a cor para o estado "vazio".
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
    }
  });
}

function handleFileImport(event) {
  const file = event.target.files[0];
  if (!file) return;

  // ✅ ALTERADO: Armazena o nome completo do arquivo TXT e atualiza o display.
  importedTxtFileName = file.name;
  updateImportedFilenamesDisplay();

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

/**
 * ✅ NOVO: Atualiza o título da "Tabela de Flash" para exibir os nomes dos arquivos importados.
 */
function updateImportedFilenamesDisplay() {
  if (!importedFilenameDisplay) return;

  const parts = [];
  if (importedExcelFileName) {
    parts.push(importedExcelFileName);
  }
  if (importedTxtFileName) {
    parts.push(importedTxtFileName);
  }

  if (parts.length > 0) {
    importedFilenameDisplay.textContent = ` - (${parts.join(" - ")})`;
  } else {
    importedFilenameDisplay.textContent = "";
  }
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
    Emergency: "EM1",
    Hazard: "HZD",
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

  // ✅ CORREÇÃO: Salva os dados da aba ativa antes de processar
  if (activeSheet) {
    const currentTableData = getCurrentTableData();
    allMacroData[activeSheet] = currentTableData;
  }

  // Mescla os dados lidos na estrutura fixa - APENAS para sheets que têm dados no arquivo TXT
  populatedSheets.forEach((name) => {
    const newSheetData = JSON.parse(JSON.stringify(baseSheetStructure));
    const parsedRows = tempData[name];

    for (let i = 0; i < parsedRows.length; i++) {
      if (i < newSheetData.length) {
        newSheetData[i] = parsedRows[i];
      }
    }

    // ✅ CORREÇÃO: Garante que todas as linhas tenham o número da linha sequencial
    newSheetData.forEach((row, index) => {
      row.row = index + 1;
    });

    allMacroData[name] = newSheetData;
  });

  // ✅ CORREÇÃO: Define FP1 como padrão SEMPRE após importação
  // Se já estamos na FP1, força o render dos dados
  if (activeSheet === "FP1") {
    renderTable(allMacroData["FP1"]);
    updateActiveSheetUI();
    updateConsumptionDisplay();
  } else {
    switchSheet("FP1", true);
  }

  renderAllActiveSettings();
  updateConsumptionDisplay();
  // ✅ NOVO: Valida os dados importados imediatamente
  validateAndReportImportErrors();
}

let errorModalTimeout; // Variável para controlar o delay ao esconder o modal.

/**
 * ✅ NOVO: Lida com o evento de passar o mouse sobre a aba.
 * @param {MouseEvent} e - O evento de mouseover.
 */
function handleSheetTabHover(e) {
  // Mostra o modal de erro se o mouse estiver sobre o ícone de aviso.
  if (e.target.classList.contains("warning-icon")) {
    clearTimeout(errorModalTimeout); // Cancela qualquer timeout para esconder o modal.
    const icon = e.target;
    const sheetName = icon.closest(".sheet-selector-btn").dataset.sheet;
    const sheetErrors = importErrorsBySheet[sheetName] || [];
    if (sheetErrors.length > 0) {
      showConfigurationErrorsModal(sheetErrors, icon); // Passa o ícone como referência de posição.
    }
  }
}

/**
 * ✅ NOVO: Lida com o evento de tirar o mouse da aba.
 * @param {MouseEvent} e - O evento de mouseout.
 */
function handleSheetTabMouseOut(e) {
  // Esconde o modal com um pequeno atraso para permitir que o usuário mova o mouse para dentro do modal.
  if (e.target.classList.contains("warning-icon")) {
    errorModalTimeout = setTimeout(() => {
      if (errorModal) {
        errorModal.style.display = "none";
      }
    }, 300); // Atraso de 300ms.
  }
}

/**
 * ✅ NOVO: Valida os dados recém-importados e exibe um modal com erros, se houver.
 */
function validateAndReportImportErrors() {
  const errors = [];

  // Se não houver canais não configurados, não há o que validar.
  // ✅ NOVO: Limpa os erros antigos e os ícones antes de uma nova validação.
  importErrorsBySheet = {};
  updateSheetTabsWithWarningIcons();

  if (!unconfiguredChannels || unconfiguredChannels.size === 0) {
    console.log(
      "Validação de importação pulada: Nenhum canal não configurado foi definido."
    );
    return;
  }
  console.log(
    "Iniciando validação de importação contra canais não configurados:",
    [...unconfiguredChannels]
  );
  sheetNames.forEach((sheetName) => {
    const sheetData = allMacroData[sheetName];
    const sheetErrors = []; // ✅ NOVO: Array para erros da planilha atual.

    sheetData.forEach((rowData) => {
      rowData.values.forEach((value, colIndex) => {
        const channelNumber = colIndex + 1;
        // Verifica se a célula está ativa ("1") E se o canal não está configurado
        if (value === "1" && unconfiguredChannels.has(channelNumber)) {
          const error = {
            sheet: sheetName,
            row: rowData.row,
            channel: channelNumber,
          };
          errors.push(error);
          sheetErrors.push(error); // ✅ NOVO: Adiciona ao array da planilha.
        }
      });
    });

    // ✅ NOVO: Armazena os erros encontrados para esta planilha.
    if (sheetErrors.length > 0) {
      importErrorsBySheet[sheetName] = sheetErrors;
    }
  });

  if (errors.length > 0) {
    console.error(
      "Erros de configuração encontrados na importação:",
      importErrorsBySheet
    );
    showConfigurationErrorsModal(errors);
  } else {
    console.log("Validação de importação concluída: Nenhum erro encontrado.");
  }

  // ✅ NOVO: Após a validação, atualiza a UI para mostrar os ícones.
  updateSheetTabsWithWarningIcons();
}

/**
 * ✅ NOVO: Adiciona ou remove ícones de aviso das abas das planilhas.
 */
function updateSheetTabsWithWarningIcons() {
  document.querySelectorAll(".sheet-selector-btn").forEach((btn) => {
    const sheetName = btn.dataset.sheet;
    const existingIcon = btn.querySelector(".warning-icon");
    if (existingIcon) {
      existingIcon.remove();
    }

    if (importErrorsBySheet[sheetName]?.length > 0) {
      const warningIcon = document.createElement("i");
      warningIcon.className = "fa fa-exclamation-triangle warning-icon";
      warningIcon.style.color = "#f39c12";
      warningIcon.style.marginLeft = "5px"; // Reduzido para melhor ajuste
      warningIcon.title =
        "Esta planilha contém erros de configuração. Clique para ver.";
      btn.appendChild(warningIcon);
    }
  });
}

/**
 * ✅ NOVO: Exibe o modal com a lista de erros de configuração.
 * @param {Array} errors - Um array de objetos de erro.
 * @param {HTMLElement} referenceElement - O elemento (ícone) que acionou o modal, usado para posicionamento.
 */
function showConfigurationErrorsModal(errors, referenceElement) {
  if (!errors || errors.length === 0) return;

  errorListContainer.innerHTML = ""; // Limpa a lista anterior
  const ul = document.createElement("ul");
  errors.forEach((err) => {
    ul.innerHTML += `<li><strong>Linha ${err.row}</strong> &rarr; <strong>Canal ${err.channel}</strong> está ativo indevidamente.</li>`;
  });
  errorListContainer.appendChild(ul);
  // ✅ NOVO: Atualiza o título do modal para ser mais específico.
  const sheetName = errors[0]?.sheet || "Desconhecida";
  errorModal.querySelector(
    "h3"
  ).innerHTML = `<i class="fa fa-exclamation-triangle"></i> Erros na Planilha "${sheetName}"`;
  errorModal.style.display = "flex";

  // ✅ NOVO: Posiciona o modal abaixo do elemento de referência (o ícone).
  if (referenceElement) {
    const rect = referenceElement.getBoundingClientRect();
    const modalContent = errorModal.querySelector(".modal-content");

    // Posiciona o modal 10px abaixo do ícone.
    modalContent.style.top = `${window.scrollY + rect.bottom + 10}px`;
    modalContent.style.left = `${window.scrollX + rect.left}px`;
  }

  // ✅ NOVO: Mantém o modal visível se o mouse entrar nele.
  errorModal.addEventListener("mouseenter", () => {
    clearTimeout(errorModalTimeout);
  });

  // ✅ NOVO: Esconde o modal se o mouse sair dele.
  errorModal.addEventListener("mouseleave", () => {
    errorModal.style.display = "none";
  });
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
        bg: "#0000FF",
        text: "white",
      },
      white: {
        bg: "#ffffff",
        text: "black",
      },
      green: {
        bg: "#00CC00",
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
      green: colorMap.green,
      yellow: colorMap.yellow,
    };

    try {
      // 2. LOOP DE RENDERIZAÇÃO: Itera sobre cada passo da macro
      for (let i = 0; i < macroData.length; i++) {
        const step = macroData[i];
        const values = step.values.map((v) => parseInt(v) || 0); // Limpa o canvas e desenha o fundo

        ctx.clearRect(0, 0, canvas.width, canvas.height);
        // ✅ CORREÇÃO: Preenche o fundo com branco sólido para evitar artefatos pretos.
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0, 0, canvas.width, canvas.height);

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
            // ✅ CORREÇÃO: Usa a mesma cor cinza da interface para consistência visual.
            ctx.fillStyle = "#6c6c6c";
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

    // ✅ ALTERAÇÃO: Usa o nome do arquivo TXT importado, se disponível.
    const zipFileName = importedTxtFileName
      ? `GIFs - ${importedTxtFileName}.zip`
      : `simulacoes_gifs_${new Date()
        .toLocaleDateString("pt-BR")
        .replace(/\//g, "-")}.zip`;
    link.download = zipFileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    console.log(
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

/**
 * ✅ OTIMIZAÇÃO: Aplica a cor a uma coluna inteira de forma otimizada.
 * Em vez de iterar sobre cada célula, ele adiciona/remove uma classe no cabeçalho
 * e usa seletores CSS para aplicar o estilo a todas as células da coluna de uma vez.
 */
function applyColumnColorOptimized(colIndex, color) {
  const ALL_COLOR_CLASSES = [
    "persistent-color-red",
    "persistent-color-green",
    "persistent-color-blue",
    "persistent-color-white",
    "persistent-color-yellow",
  ];
  const newColorClass = `persistent-color-${color}`;

  const headerCell = document.querySelector(
    `#macro-header th[data-col-index="${colIndex}"]`
  );
  if (!headerCell) return;

  // Limpa cores antigas do cabeçalho
  headerCell.classList.remove(...ALL_COLOR_CLASSES);

  if (color === "clear") {
    delete columnColors[colIndex];
  } else {
    columnColors[colIndex] = color;
    headerCell.classList.add(newColorClass);
  }

  // Força a atualização visual das células da coluna que já têm conteúdo
  const cellsToUpdate = macroTableBody.querySelectorAll(
    `td[data-col-index="${colIndex}"]`
  );
  cellsToUpdate.forEach(updateCellColor);
}

function applyAllColumnColors() {
  for (const colIndex in columnColors) {
    applyColumnColorOptimized(parseInt(colIndex, 10), columnColors[colIndex]);
  }
}

// === INICIALIZAÇÃO ===
// === INICIALIZAÇÃO ===
document.addEventListener("DOMContentLoaded", initializeApp);
