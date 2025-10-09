<div class="right-panel-container">
    <div class="center-panel">
        <div
            style="
              width: 100%;
              display: flex;
              justify-content: space-between;
              align-items: center;
              flex-wrap: wrap;
              gap: 10px;
            ">
            <div
                class="panel-title"
                style="border-bottom: none; margin-bottom: 0">
                Barra:
                <span id="current-layout-name" style="font-weight: 500"></span>
            </div>
            <div style="display: flex; gap: 10px">
                <div style="display: flex; gap: 10px">
                    <button
                        id="open-layout-panel-btn"
                        class="btn"
                        style="
                    background-color: #16a085;
                    font-size: 10px;
                    padding: 4px 8px;
                  ">
                        <i class="fa fa-th"></i> Selecionar Modelo
                    </button>
                    <button
                        id="edit-layout-btn"
                        class="btn"
                        style="
                    background-color: #3498db;
                    font-size: 10px;
                    padding: 4px 8px;
                  ">
                        <i class="fa fa-pencil"></i> Editar Layout
                    </button>
                    <button
                        id="save-layout-btn"
                        class="btn"
                        style="
                    background-color: #27ae60;
                    font-size: 10px;
                    padding: 4px 8px;
                    display: none;
                  ">
                        <i class="fa fa-save"></i> Salvar Layout
                    </button>
                    <button
                        id="export-layout-btn"
                        class="btn"
                        style="
                    background-color: #9b59b6;
                    font-size: 10px;
                    padding: 4px 8px;
                    display: none;
                  ">
                        <i class="fa fa-code"></i> Exportar Coordenadas
                    </button>
                </div>
            </div>
            <div class="simulation-area" id="simulation-area">
                <img src="" alt="Modelo da Sirene" id="siren-model-image" />
                <div id="simulation-placeholder" class="simulation-placeholder">
                    <i class="fa fa-th-large"></i>
                    <p>Selecione um modelo para começar</p>
                    <span>Use o botão "Selecionar Modelo" ou "Extrair Layout Excel"
                        para carregar uma barra.</span>
                </div>
                <div id="blocks-container"></div>
            </div>
        </div>
        <div class="right-panel">

            <div class="panel-title">Botões / Funções</div>
            <div class="config-image-container row" style="margin: 15px auto; min-height: 150px;">
                <div class="column d-flex flex-column" style="gap: 10px; max-height: 120px;">
                    <div class="d-flex justify-content-around" style="left: 5%; position: relative;">
                        <button onclick="triggerMacro('FP1')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">FP1</button>
                        <button onclick="triggerMacro('FP2')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">FP2</button>
                        <button onclick="triggerMacro('FP3')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">FP3</button>
                        <button onclick="triggerMacro('FP4')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">FP4</button>
                    </div>
                    <div class="d-flex justify-content-around" style="left: 5%; position: relative;">
                        <button onclick="triggerMacro('AUX1')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">AUX1</button>
                        <button onclick="triggerMacro('AUX2')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">AUX2</button>
                        <button onclick="triggerMacro('AUX3')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">AUX3</button>
                        <button onclick="triggerMacro('AUX4')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">AUX4</button>
                    </div>
                    <div class="d-flex justify-content-around">
                        <button onclick="triggerMacro('DS-L')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">DS-L</button>
                        <button onclick="triggerMacro('DS-C')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">DS-C</button>
                        <button onclick="triggerMacro('DS-R')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">DS-R</button>
                        <button onclick="triggerMacro('EM1')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">EM1</button>
                        <button onclick="triggerMacro('HZD')" class="btn btn-sm" style="flex: 1; margin: 2px; font-size: 11px; width:10%;">HZD</button>
                    </div>
                    <div class="column d-flex flex-column" style="gap: 10px; margin-top: 10px; left: 70%; position: relative; top: -110px;">
                        <button id="btn-le" onclick="toggleSpecialFunction('alley_left')" class="control-btn btn btn-sm" style="font-size: 11px; flex: 1; height: 100px;">LE</button>
                        <button id="btn-td" onclick="toggleSpecialFunction('takedown')" class="control-btn btn btn-sm" style="font-size: 11px; flex: 1; height: 100px;">TD</button>
                        <button id="btn-ld" onclick="toggleSpecialFunction('alley_right')" class="control-btn btn btn-sm" style="font-size: 11px; flex: 1; height: 100px;">LD</button>
                    </div>
                </div>
                <div class="output-item" id="special-functions-status-output">
                    <div class="output-label">Funções Especiais Ativas</div>
                    <div id="special-functions-active-list" class="status-panel-list">
                        <p class="no-settings">Nenhuma função especial ativa.</p>
                    </div>
                </div>
                <div class="output-item" id="jumptabela-status-output">
                    <div class="output-label">
                        Configurações Globais Ativas (JumpTabela)
                    </div>
                    <div id="jumptabela-active-list" class="status-panel-list">
                        <p class="no-settings">Nenhuma configuração global ativa.</p>
                    </div>
                </div>
                <div class="output-item" style="background-color: #e9ecef">
                    <div class="output-label" style="color: #2c3e50; font-weight: 700">
                        CONSUMO DO MÓDULO
                    </div>
                    <div style="
                display: flex;
                justify-content: space-around;
                margin-top: 5px;
              ">
                        <div style="text-align: center">
                            <div class="output-label">MÉDIA FINAL</div>
                            <div class="output-value" id="output-media-final"
                                style="font-size: 14px; font-weight: bold; color: #0070c0">
                                0,000 A
                            </div>
                        </div>
                        <div style="text-align: center">
                            <div class="output-label">PICO</div>
                            <div class="output-value" id="output-pico"
                                style="font-size: 14px; font-weight: bold; color: #dc3545">
                                0,000 A
                            </div>
                        </div>
                    </div>
                </div>
                <button class="btn" id="edit-config-btn" style="width: 100%; margin-top: 10px">
                    Editar Configuração
                </button>

                <!-- ✅ NOVO: Modal de Erros de Configuração -->
                <div id="error-modal" class="modal-overlay">
                    <div class="modal-content" style="width: 500px;">
                        <div class="modal-header">
                            <h3 style="color: #c0392b;"><i class="fa fa-exclamation-triangle"></i> Erros de Configuração na Importação</h3>
                            <button id="close-error-modal-btn" class="close-button">&times;</button>
                        </div>
                        <div class="modal-body">
                            <p>Foram encontrados os seguintes problemas ao importar o arquivo. Um canal não configurado (cinza) está sendo ativado:</p>
                            <div id="error-list-container" class="error-list-container">
                                <!-- Erros serão inseridos aqui pelo JavaScript -->
                            </div>
                        </div>
                    </div>
                </div>

            </div>
        </div>