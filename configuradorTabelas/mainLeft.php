<div class="panel left-panel">
    <div class="panel-title">
        Tabela de Flash
        <span id="imported-filename-display" style="font-weight: 400; color: #6c757d; font-size: 12px;"></span>
    </div>
    <p class="instruction-text">
        Selecione uma ou mais colunas pelo cabeçalho para definir uma cor.
    </p>
    <!-- ✅ NOVO: A paleta de cores agora é um popup que fica escondido por padrão -->
    <div id="color-palette-popup" class="color-palette-popup">
        <div class="color-selector" id="color-palette">
            <div class="color-swatch" data-color="red" style="background-color: #FF0000;" title="Vermelho"></div>
            <div class="color-swatch" data-color="blue" style="background-color: #0000FF;" title="Azul"></div>
            <div class="color-swatch" data-color="white" style="background-color: #FFFFFF;" title="Branco"></div>
            <div class="color-swatch" data-color="green" style="background-color: #00CC00;" title="Verde"></div>
            <div class="color-swatch" data-color="yellow" style="background-color: #FFC000;" title="Âmbar"></div>
            <div class="color-swatch clear-swatch" data-color="clear" title="Limpar Cor"><i class="fa fa-times"></i></div>
        </div>
    </div>
    <div class="sheet-selector-container" id="sheet-selector-container">
        <!-- As abas das planilhas são geradas dinamicamente pelo JavaScript -->
    </div>
    <div class="excel-container">
        <div class="excel-toolbar">
            <button
                id="gravar-btn"
                class="btn"
                style="
                background-color: #c40303;
                font-size: 9px;
                padding: 3px 8px;
              ">
                Gravar
            </button>
            <button
                id="import-txt-btn"
                class="btn"
                style="
                background-color: #9b59b6;
                font-size: 9px;
                padding: 3px 8px;
              ">
                <i class="fa fa-upload"></i> Importar TXT
            </button>
            <button
                id="extract-excel-layout-btn"
                class="btn"
                style="
                background-color: #16a085;
                font-size: 9px;
                padding: 3px 8px;
              ">
                <i class="fa fa-file-excel"></i> Extrair Layout
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
                font-size: 9px;
                padding: 3px 8px;
              ">
                <i class="fa fa-file-image"></i> Exportar GIFs
            </button>
            <button
                id="restaurar-btn"
                class="btn"
                style="
                background-color: #2ecc71;
                font-size: 9px;
                padding: 3px 8px;
              ">
                Restaurar
            </button>
            <button
                id="limpa-btn"
                class="btn"
                style="
                background-color: #f39c12;
                font-size: 9px;
                padding: 3px 8px;
              ">
                Limpar Planilha
            </button>
            <button
                id="advanced-settings-btn"
                class="btn"
                style="
                background-color: #34495e;
                font-size: 9px;
                padding: 3px 8px;
                margin-left: 15px;
              ">
                <i class="fa fa-cogs"></i> Definições
            </button>
            <button class="btn" id="execute-macro-btn" style="margin-top:5px;">
                SIMULAR
            </button>
        </div>
        <div class="excel-grid">
            <table class="macro-table">
                <thead id="macro-header"></thead>
                <tbody id="macro-body"></tbody>
            </table>
        </div>
    </div>
</div>