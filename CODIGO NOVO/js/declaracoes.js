let configModal,
  openModalBtn,
  closeModalBtn,
  saveAndCloseModalBtn,
  executeMacroBtn,
  blocksContainer,
  macroTableBody,
  macroHeader,
  sheetSelectorContainer,
  gravarBtn,
  restaurarBtn,
  limpaBtn,
  colorContextMenu,
  layoutPanelModal,
  openLayoutPanelBtn,
  closeLayoutPanelBtn,
  layoutPanelBody,
  currentLayoutNameSpan,
  advancedSettingsModal,
  openAdvancedSettingsBtn,
  closeAdvancedSettingsBtn,
  saveAdvancedSettingsBtn,
  tabButtons,
  tabContents,
  jumpTabelaCheckboxes,
  specialFunctionCells,
  jumpTabelaStatusContainer,
  specialFunctionsStatusContainer,
  simulationArea,
  editLayoutBtn,
  saveLayoutBtn;

// === GERENCIAMENTO DE ESTADO ===
let simulationStopper = null;
let isSimulationRunning = false;
const sheetNames = [
  "FP1",
  "FP2",
  "FP3",
  "FP4",
  "AUX1",
  "AUX2",
  "AUX3",
  "AUX4",
  "DS-L",
  "DS-C",
  "DS-R",
  "HZD",
  "EM1",
];
let allMacroData = {};
let activeSheet = "FP1";
let savedSpecialFunctionConfigs = {};
let savedJumpTabelaSettings = {};
let columnColors = {};
let clipboardData = null;
let copiedBlockData = null;
let isEditMode = false;
let activeBlock = null;
let action = null;
let activeSpecialFunctions = {};
const selection = new Set();
let undoHistory = {};
let redoHistory = {};
let debounceTimer;
let gifWorkerBlobURL = null;
let backgroundImageDataURL = null;
let currentLayout;
let activeBlockLayout;
let allAvailableLayouts = []; // Variável global para a lista de layouts
let selectedColumns = new Set();
let lastSelectedColumn = -1;

// ================================================================================= //
// ======================= MAPA DE CÓDIGOS DE ARQUIVO PARA LAYOUT ====================== //
// ================================================================================= //

const fileCodeToLayoutMap = {
  "00179": "primus47p18mCol",
  "00180": "primus47p20mRef",
  "00181": "primus40p18mCol",
  "00182": "primus47p20mRef",
  "00183": "primus47p20mRef",
  "00184": "primus47p20mRef",
  "00185": "primus47p20mRef",
  "00186": "venus11m",
  "00187": "primus47p20mRef",
  "00188": "primus47p20mRef",
  "00189": "primus47p18mRef",
  "00190": "primus47p20mRef",
  "00191": "venus14m",
  "00193": "venus24m",
  "00194": "primus47p24mRef",
  "00195": "venus14m",
  "00196": "primus61p24mRef",
  "00197": "primus47p13mRef",
  "00198": "venus22m",
  "00199": "primus54p22mRef",
  "00200": "primus47p11mCol",
  "00201": "primus47p16mCol",
  "00202": "primus47p18mRef",
  "00203": "primus47p16mRef",
  "00204": "primus47p20mRef",
  "00205": "primus47p21mRef",
  "00207": "primus47p20mRef",
  "00208": "primus47p22mRef",
  "00209": "primus47p22mRef",
  "00210": "primus47p24mRef",
  "00211": "primus40p18mRef",
  "00212": "primus54p18mCol",
  "00213": "primus47p20mRef",
  "00214": "primus47p16mCol",
  "00215": "primus54p22mRef",
  "00216": "primus47p16mCol",
  "00217": "venus15m",
  "00218": "primus47p16mCol",
  "00219": "primus47p20mCol",
  "00220": "primus40p18mRef",
  "00221": "primus47p20mRef",
  "00222": "venus15m",
  "00223": "primus47p20mRef",
  "00224": "primus47p20mRef",
  "00225": "primus47p16mCol",
  "00226": "venus15m",
  "00227": "primus40p14mCol",
  "00228": "primus47p16mRef",
  "00229": "primus54p22mRef",
  "00232": "venus15m",
  "00233": "primus47p20mRef",
  "00234": "primus47p24mCol",
  "00235": "primus47p13mRef",
};

const layouts = [
  {
    id: "s_slim_24_default",
    name: "S-SLIM 24m Padrão",
    image: "simulação.png",
    category: "Padrão",
    blocks: s_slim_24_default_BlockLayout,
  },
  {
    id: "s_slim_bicolor_example",
    name: "S-SLIM Bicolor Exemplo",
    image: "simulação.png",
    category: "Padrão",
    blocks: s_slim_bicolor_example,
  },

  // 40 Polegadas
  {
    id: "primus40p12mCol",
    name: "Primus 40p 12m Col",
    image: "ImagensSirene40p/primus40p12mCol.png",
    category: "40 Polegadas",
    blocks: primus40p12mCol_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus40p12mRef",
    name: "Primus 40p 12m Ref",
    image: "ImagensSirene40p/primus40p12mRef.png",
    category: "40 Polegadas",
    blocks: primus40p12mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus40p14mCol",
    name: "Primus 40p 14m Col",
    image: "ImagensSirene40p/primus40p14mCol.png",
    category: "40 Polegadas",
    blocks: primus40p14mCol_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus40p14mRef",
    name: "Primus 40p 14m Ref",
    image: "ImagensSirene40p/primus40p14mRef.png",
    category: "40 Polegadas",
    blocks: primus40p14mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus40p16mCol",
    name: "Primus 40p 16m Col",
    image: "ImagensSirene40p/primus40p16mCol.png",
    category: "40 Polegadas",
    blocks: primus40p16mCol_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus40p16mRef",
    name: "Primus 40p 16m Ref",
    image: "ImagensSirene40p/primus40p16mRef.png",
    category: "40 Polegadas",
    blocks: primus40p16mRef_BlockLayout,
  },
  {
    id: "primus40p18mCol",
    name: "Primus 40p 18m Col",
    image: "ImagensSirene40p/primus40p18mCol.png",
    category: "40 Polegadas",
    blocks: primus40p18mCol_BlockLayout,
  },
  {
    id: "primus40p18mRef",
    name: "Primus 40p 18m Ref",
    image: "ImagensSirene40p/primus40p18mRef.png",
    category: "40 Polegadas",
    blocks: primus40p18mRef_BlockLayout,
  },
  {
    id: "primus40p22mRef",
    name: "Primus 40p 22m Ref",
    image: "ImagensSirene40p/primus40p22mRef.png",
    category: "40 Polegadas",
    blocks: primus40p22mRef_BlockLayout,
  },

  // 47 Polegadas
  {
    id: "primus47p11mCol",
    name: "Primus 47p 11m Col",
    image: "ImagensRefletor47p/colimador/primus47p11mCol.png",
    category: "47 Polegadas",
    blocks: primus47p11mCol_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p11mRef",
    name: "Primus 47p 11m Ref",
    image: "ImagensRefletor47p/refletor/primus47p11mRef.png",
    category: "47 Polegadas",
    blocks: primus47p11mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p13mRef",
    name: "Primus 47p 13m Ref",
    image: "ImagensRefletor47p/refletor/primus47p13mRef.png",
    category: "47 Polegadas",
    blocks: primus47p13mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p14mCol",
    name: "Primus 47p 14m Col",
    image: "ImagensRefletor47p/colimador/primus47p14mCol.png",
    category: "47 Polegadas",
    blocks: primus47p14mCol_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p14mRef",
    name: "Primus 47p 14m Ref",
    image: "ImagensRefletor47p/refletor/primus47p14mRef.png",
    category: "47 Polegadas",
    blocks: primus47p14mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p16mCol",
    name: "Primus 47p 16m Col",
    image: "ImagensRefletor47p/colimador/primus47p16mCol.png",
    category: "47 Polegadas",
    blocks: primus47p16mCol_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p16mRef",
    name: "Primus 47p 16m Ref",
    image: "ImagensRefletor47p/refletor/primus47p16mRef.png",
    category: "47 Polegadas",
    blocks: primus47p16mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p18mRef",
    name: "Primus 47p 18m Ref",
    image: "ImagensRefletor47p/refletor/primus47p18mRef.png",
    category: "47 Polegadas",
    blocks: primus47p18mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p20mCol",
    name: "Primus 47p 20m Col",
    image: "ImagensRefletor47p/colimador/primus47p20mCol.png",
    category: "47 Polegadas",
    blocks: primus47p20mCol_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p20mRef",
    name: "Primus 47p 20m Ref",
    image: "ImagensRefletor47p/refletor/primus47p20mRef.png",
    category: "47 Polegadas",
    blocks: primus47p20mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p22mRef",
    name: "Primus 47p 22m Ref",
    image: "ImagensRefletor47p/refletor/primus47p22mRef.png",
    category: "47 Polegadas",
    blocks: primus47p22mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico
  {
    id: "primus47p24mRef",
    name: "Primus 47p 24m Ref",
    image: "ImagensRefletor47p/refletor/primus47p24mRef.png",
    category: "47 Polegadas",
    blocks: primus47p24mRef_BlockLayout,
  }, // <-- Manter padrão ou criar um específico

  /// 54 Polegadas
  {
    id: "primus54p10mCol",
    name: "Primus 54p 10m Col",
    image: "ImagensSirene54P/colimador/primus54p10mCol.png",
    category: "54 Polegadas",
    blocks: primus54p10mCol_BlockLayout,
  },
  {
    id: "primus54p12mRef",
    name: "Primus 54p 12m Ref",
    image: "ImagensSirene54P/refletor/primus54p12mRef.png",
    category: "54 Polegadas",
    blocks: primus54p12mRef_BlockLayout,
  },
  {
    id: "primus54p14mCol",
    name: "Primus 54p 14m Col",
    image: "ImagensSirene54P/colimador/primus54p14mCol.png",
    category: "54 Polegadas",
    blocks: primus54p14mCol_BlockLayout,
  },
  {
    id: "primus54p14mRef",
    name: "Primus 54p 14m Ref",
    image: "ImagensSirene54P/refletor/primus54p14mRef.png",
    category: "54 Polegadas",
    blocks: primus54p14mRef_BlockLayout,
  },
  {
    id: "primus54p16mCol",
    name: "Primus 54p 16m Col",
    image: "ImagensSirene54P/colimador/primus54p16mCol.png",
    category: "54 Polegadas",
    blocks: primus54p16mCol_BlockLayout,
  },
  {
    id: "primus54p18mRef",
    name: "Primus 54p 18m Ref",
    image: "ImagensSirene54P/refletor/primus54p18mRef.png",
    category: "54 Polegadas",
    blocks: primus54p18mRef_BlockLayout,
  },
  {
    id: "primus54p20mRef",
    name: "Primus 54p 20m Ref",
    image: "ImagensSirene54P/refletor/primus54p20mRef.png",
    category: "54 Polegadas",
    blocks: primus54p20mRef_BlockLayout,
  },
  {
    id: "primus54p22mCol",
    name: "Primus 54p 22m Col",
    image: "ImagensSirene54P/colimador/primus54p22mCol.png",
    category: "54 Polegadas",
    blocks: primus54p22mCol_BlockLayout,
  },
  {
    id: "primus54p22mRef",
    name: "Primus 54p 22m Ref",
    image: "ImagensSirene54P/refletor/primus54p22mRef.png",
    category: "54 Polegadas",
    blocks: primus54p22mRef_BlockLayout,
  },
  {
    id: "primus54p26mRef",
    name: "Primus 54p 26m Ref",
    image: "ImagensSirene54P/refletor/primus54p26mRef.png",
    category: "54 Polegadas",
    blocks: primus54p26mRef_BlockLayout,
  },

  // 61 Polegadas
  {
    id: "primus61p13mCol",
    name: "Primus 61p 13m Col",
    image: "ImagensSirene61P/colimador/primus61p13mCol.png",
    category: "61 Polegadas",
    blocks: primus61p13mCol_BlockLayout,
  },
  {
    id: "primus61p13mRef",
    name: "Primus 61p 13m Ref",
    image: "ImagensSirene61P/refletor/primus61p13mRef.png",
    category: "61 Polegadas",
    blocks: primus61p13mRef_BlockLayout,
  },
  {
    id: "primus61p14mCol",
    name: "Primus 61p 14m Col",
    image: "ImagensSirene61P/colimador/primus61p14mCol.png",
    category: "61 Polegadas",
    blocks: s_slim_24_default_BlockLayout,
  },
  {
    id: "primus61p15mRef",
    name: "Primus 61p 15m Ref",
    image: "ImagensSirene61P/refletor/primus61p15mRef.png",
    category: "61 Polegadas",
    blocks: primus61p15mRef_BlockLayout,
  },
  {
    id: "primus61p18mCol",
    name: "Primus 61p 18m Col",
    image: "ImagensSirene61P/colimador/primus61p18mCol.png",
    category: "61 Polegadas",
    blocks: primus61p18mCol_BlockLayout,
  },
  {
    id: "primus61p18mRef",
    name: "Primus 61p 18m Ref",
    image: "ImagensSirene61P/refletor/primus61p18mRef.png",
    category: "61 Polegadas",
    blocks: primus61p18mRef_BlockLayout,
  },
  {
    id: "primus61p20mRef",
    name: "Primus 61p 20m Ref",
    image: "ImagensSirene61P/refletor/primus61p20mRef.png",
    category: "61 Polegadas",
    blocks: primus61p20mRef_BlockLayout,
  },
  {
    id: "primus61p22mRef",
    name: "Primus 61p 22m Ref",
    image: "ImagensSirene61P/refletor/primus61p22mRef.png",
    category: "61 Polegadas",
    blocks: primus61p22mRef_BlockLayout,
  },
  {
    id: "primus61p24mCol",
    name: "Primus 61p 24m Col",
    image: "ImagensSirene61P/colimador/primus61p24mCol.png",
    category: "61 Polegadas",
    blocks: primus61p24mCol_BlockLayout,
  },
  {
    id: "primus61p24mRef",
    name: "Primus 61p 24m Ref",
    image: "ImagensSirene61P/refletor/primus61p24mRef.png",
    category: "61 Polegadas",
    blocks: primus61p24mRef_BlockLayout,
  },
  {
    id: "primus61p28mRef",
    name: "Primus 61p 28m Ref",
    image: "ImagensSirene61P/refletor/primus61p28mRef.png",
    category: "61 Polegadas",
    blocks: primus61p28mRef_BlockLayout,
  },

  // Venus
  {
    id: "venus11m",
    name: "Venus 11m",
    image: "VenusSirene/venus11m.png",
    category: "Venus",
    blocks: s_slim_24_default_BlockLayout,
  },
  {
    id: "venus14m",
    name: "Venus 14m",
    image: "VenusSirene/venus14m.png",
    category: "Venus",
    blocks: s_slim_24_default_BlockLayout,
  },
  {
    id: "venus15m",
    name: "Venus 15m",
    image: "VenusSirene/venus15m.png",
    category: "Venus",
    blocks: s_slim_24_default_BlockLayout,
  },
  {
    id: "venus16m",
    name: "Venus 16m",
    image: "VenusSirene/venus16m.png",
    category: "Venus",
    blocks: s_slim_24_default_BlockLayout,
  },
  {
    id: "venus19m",
    name: "Venus 19m",
    image: "VenusSirene/venus19m.png",
    category: "Venus",
    blocks: s_slim_24_default_BlockLayout,
  },
  {
    id: "venus20m",
    name: "Venus 20m",
    image: "VenusSirene/venus20m.png",
    category: "Venus",
    blocks: s_slim_24_default_BlockLayout,
  },
  {
    id: "venus22m",
    name: "Venus 22m",
    image: "VenusSirene/venus22m.png",
    category: "Venus",
    blocks: s_slim_24_default_BlockLayout,
  },
  {
    id: "venus24m",
    name: "Venus 24m",
    image: "VenusSirene/venus24m.png",
    category: "Venus",
    blocks: s_slim_24_default_BlockLayout,
  },
];
