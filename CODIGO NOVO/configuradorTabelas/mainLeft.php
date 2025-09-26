<div class="panel left-panel">
    <div class="panel-title">Controle de Macro</div>
    <p class="instruction-text">
        Clique com o botão direito no número da coluna (1-24) para definir uma
        cor fixa para o canal.
    </p>
    <div
        id="sheet-selector-container"
        class="sheet-selector-container"></div>
    <div class="excel-container">
        <div class="excel-toolbar">
            <button
                id="gravar-btn"
                class="btn"
                style="
                background-color: #c40303;
                font-size: 10px;
                padding: 3px 8px;
              ">
                Gravar
            </button>
            <button
                id="import-txt-btn"
                class="btn"
                style="
                background-color: #9b59b6;
                font-size: 10px;
                padding: 3px 8px;
              ">
                <i class="fa fa-upload"></i> Importar TXT
            </button>
            <button
                id="extract-excel-layout-btn"
                class="btn"
                style="
                background-color: #16a085;
                font-size: 10px;
                padding: 3px 8px;
              ">
                <i class="fa fa-file-excel"></i> Extrair Layout Excel
            </button>
            <input
                type="file"
                id="excel-layout-input"
                style="display: none"
                accept=".xlsx, .xlsm" />
            <button
                id="export-gif-btn"
                class="btn"
                style="
                background-color: #1abc9c;
                font-size: 10px;
                padding: 3px 8px;
              ">
                <i class="fa fa-file-image"></i> Exportar GIFs
            </button>
            <button
                id="restaurar-btn"
                class="btn"
                style="
                background-color: #2ecc71;
                font-size: 10px;
                padding: 3px 8px;
              ">
                Restaurar
            </button>
            <button
                id="limpa-btn"
                class="btn"
                style="
                background-color: #f39c12;
                font-size: 10px;
                padding: 3px 8px;
              ">
                Limpar Planilha
            </button>
            <button
                id="advanced-settings-btn"
                class="btn"
                style="
                background-color: #34495e;
                font-size: 10px;
                padding: 3px 8px;
                margin-left: 15px;
              ">
                <i class="fa fa-cogs"></i> Definições
            </button>
        </div>
        <div class="excel-grid">
            <table class="macro-table">
                <thead id="macro-header"></thead>
                <tbody id="macro-body"></tbody>
            </table>
        </div>
    </div>
    <div class="macro-controls">
        <button class="btn" id="execute-macro-btn">Executar Macro</button>
    </div>
</div>