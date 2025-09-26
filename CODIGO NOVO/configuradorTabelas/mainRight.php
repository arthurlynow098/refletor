<div class="right-panel-container">
    <div class="center-panel">
        <div
            style="
                margin-top: 100px;
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
                Simulação:
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
            <div style="width: 831px; margin: 0 auto;"> <!-- Container com largura fixa -->
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
        </div>
        <div class="right-panel">


            <div class="panel-title">Resultado da Configuração</div>
            <div
                class="config-image-container"
                style="
              position: relative;
              width: 100%;
              max-width: 500px;
              margin: 15px auto;
            ">
                <img
                    src="Imagemcontrole.png"
                    alt="Controle da Barra"
                    style="
                width: 100%;
                height: auto;
                display: block;
                border-radius: 15px;
                opacity: 0.5;
              " />
                <button
                    onclick="triggerMacro('FP1')"
                    style="
                position: absolute;
                top: 20%;
                left: 16%;
                width: 15%;
                height: 12%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
              ">
                    FP1
                </button>
                <button
                    onclick="triggerMacro('FP3')"
                    style="
                position: absolute;
                top: 20%;
                left: 42.5%;
                width: 15%;
                height: 12%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
              ">
                    FP3
                </button>
                <button
                    onclick="triggerMacro('FP2')"
                    style="
                position: absolute;
                top: 20%;
                left: 69%;
                width: 15%;
                height: 12%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
              ">
                    FP2
                </button>
                <button
                    id="btn-le"
                    onclick="toggleSpecialFunction('alley_left')"
                    class="control-btn"
                    style="
                position: absolute;
                top: 26%;
                left: 6%;
                width: 6%;
                height: 50%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: all 0.2s ease;
              ">
                    LE
                </button>
                <button
                    id="btn-ld"
                    onclick="toggleSpecialFunction('alley_right')"
                    class="control-btn"
                    style="
                position: absolute;
                top: 26%;
                left: 88%;
                width: 6%;
                height: 50%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: all 0.2s ease;
              ">
                    LD
                </button>
                <button
                    id="btn-td"
                    onclick="toggleSpecialFunction('takedown')"
                    class="control-btn"
                    style="
                position: absolute;
                top: 51%;
                left: 42.5%;
                width: 15%;
                height: 12%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: all 0.2s ease;
              ">
                    TD
                </button>
                <button
                    onclick="triggerMacro('EM1')"
                    style="
                position: absolute;
                top: 69%;
                left: 15.4%;
                width: 15%;
                height: 12%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
              ">
                    EM1
                </button>
                <button
                    onclick="triggerMacro('HZD')"
                    style="
                position: absolute;
                top: 69%;
                left: 70%;
                width: 15%;
                height: 12%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
              ">
                    HZD
                </button>
                <button
                    onclick="triggerMacro('DS-L')"
                    style="
                position: absolute;
                top: 70%;
                left: 33%;
                width: 10%;
                height: 10%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
              ">
                    «
                </button>
                <button
                    onclick="triggerMacro('DS-C')"
                    style="
                position: absolute;
                top: 75%;
                left: 45%;
                width: 10%;
                height: 10%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
              ">
                    «»
                </button>
                <button
                    onclick="triggerMacro('DS-R')"
                    style="
                position: absolute;
                top: 70%;
                left: 57%;
                width: 10%;
                height: 10%;
                background-color: #e9ecef;
                border: 1px solid #555;
                cursor: pointer;
                border-radius: 5px;
                color: #000;
                font-weight: bold;
                font-size: 1vw;
                display: flex;
                align-items: center;
                justify-content: center;
              ">
                    »
                </button>
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
                <div
                    style="
                display: flex;
                justify-content: space-around;
                margin-top: 5px;
              ">
                    <div style="text-align: center">
                        <div class="output-label">MÉDIA FINAL</div>
                        <div
                            class="output-value"
                            id="output-media-final"
                            style="font-size: 14px; font-weight: bold; color: #0070c0">
                            0,000 A
                        </div>
                    </div>
                    <div style="text-align: center">
                        <div class="output-label">PICO</div>
                        <div
                            class="output-value"
                            id="output-pico"
                            style="font-size: 14px; font-weight: bold; color: #dc3545">
                            0,000 A
                        </div>
                    </div>
                </div>
            </div>
            <button
                class="btn"
                id="edit-config-btn"
                style="width: 100%; margin-top: 10px">
                Editar Configuração
            </button>
        </div>
    </div>