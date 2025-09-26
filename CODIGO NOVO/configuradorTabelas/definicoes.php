<div id="advanced-settings-modal" class="modal-overlay">
    <div class="modal-content" style="width: 80%; max-width: 1200px">
        <div class="modal-header">
            <div
                class="panel-title"
                style="margin-bottom: 0; border-bottom: none">
                <i class="fa fa-cogs"></i> Definições Avançadas de Funcionamento
            </div>
            <button class="close-button" id="close-advanced-modal-btn">
                &times;
            </button>
        </div>
        <div class="modal-body">
            <div class="modal-nav-tabs">
                <button
                    class="modal-tab-btn active"
                    data-tab="tab-special-functions">
                    Funções Especiais
                </button>
                <button class="modal-tab-btn" data-tab="tab-jumptabela">
                    JumpTabela
                </button>
            </div>
            <div class="modal-tab-content-container">
                <div id="tab-special-functions" class="modal-tab-content active">
                    <p class="description">
                        Configure quais blocos devem acender ou apagar para cada função
                        especial. Para ativar um bloco para uma função, digite "1" na
                        coluna correspondente ao número do bloco.
                    </p>
                    <div class="function-group">
                        <h3 class="function-group-title">Take Down (TD)</h3>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Apagar Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="takedown" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Acender Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="takedown_on" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="function-group">
                        <h3 class="function-group-title">Alley-Right (LD)</h3>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Apagar Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="alley_right" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Acender Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="alley_right_on" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="function-group">
                        <h3 class="function-group-title">Alley-Left (LE)</h3>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Apagar Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="alley_left" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Acender Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="alley_left_on" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="function-group">
                        <h3 class="function-group-title">Backlight Activate</h3>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Apagar Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="backlight_off" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Acender Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="backlight_on" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="function-group">
                        <h3 class="function-group-title">Cut Front</h3>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Apagar Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="cut_front_off" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Acender Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="cut_front_on" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="function-group">
                        <h3 class="function-group-title">Cut Rear</h3>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Apagar Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="cut_rear_off" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Acender Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="cut_rear_on" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="function-group">
                        <h3 class="function-group-title">
                            Cut Direcionador de Trânsito
                        </h3>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Apagar Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="cut_dt_off" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Acender Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="cut_dt_on" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="function-group">
                        <h3 class="function-group-title">Cut Horn Light</h3>
                        <div class="function-table-container">
                            <table class="function-table">
                                <thead>
                                    <tr>
                                        <th>Instrução</th>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(`<th>${i}</th>`);
                                        </script>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Apagar Blocos</td>
                                        <script>
                                            for (let i = 1; i <= 24; i++)
                                                document.write(
                                                    `<td contenteditable="true" data-function="cut_horn_light" data-block="${i}"></td>`
                                                );
                                        </script>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="function-group">
                        <h3 class="function-group-title">Número de FP principais</h3>
                        <div class="form-group" style="margin-bottom: 15px">
                            <label class="form-label" style="text-transform: none">Informe quantos FP poderão ser usados para SL1-4 e Aux1-5
                                >></label>
                            <select class="form-select" id="config-num-fp-principais">
                                <script>
                                    for (let i = 1; i <= 15; i++)
                                        document.write(
                                            `<option value="${i}" ${
                            i === 4 ? "selected" : ""
                          }>${i}</option>`
                                        );
                                </script>
                            </select>
                            <span class="fp-setting-description">Quantidade de FP Principais precisa ser preenchido quando
                                utilizar FP de Horn. Default = 4</span>
                        </div>
                    </div>
                    <div class="function-group">
                        <h3 class="function-group-title">Quantidade de FP para DS</h3>
                        <div class="form-group">
                            <label class="form-label" style="text-transform: none">Informe quantos FP poderão ser usados para DS >></label>
                            <select class="form-select" id="config-qtd-fp-ds">
                                <script>
                                    for (let i = 1; i <= 15; i++)
                                        document.write(
                                            `<option value="${i}" ${
                            i === 8 ? "selected" : ""
                          }>${i}</option>`
                                        );
                                </script>
                            </select>
                            <span
                                class="fp-setting-description"
                                style="font-style: normal">&lt;&lt;O número deve ser multiplo de 4 pois servirá para
                                DR, DL, DC e Hazard.
                                <em>Quantidade de FP para DS precisa ser preenchido quando
                                    utilizar FP de Horn. Default = 8</em></span>
                        </div>
                    </div>
                </div>
                <div id="tab-jumptabela" class="modal-tab-content">
                    <p class="description">
                        Configure o comportamento global do sistema através dos bits da
                        JumpTabela. Marque as caixas para ativar cada função
                        (equivalente a colocar o valor "1").
                    </p>
                    <div class="jumptabela-section">
                        <h4 class="function-title">
                            JumpTabela #1 - Configurações Gerais
                        </h4>
                        <div class="checkbox-grid">
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="1" data-bit="7" />
                                b7 >> Ativa o modo fio nos moldes da PCESP</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="1" data-bit="6" />
                                b6 >> Remove PWM nos canais 21 e 24 para usar Bloco de LED
                                ou RELÉ</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="1" data-bit="5" />
                                b5 >> Define que o MB funciona em sincronismo com outras MB
                                e/ou Auxiliares</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="1" data-bit="4" />
                                b4 >> Aciona a saída Alley-Left sem espera do sincronismo
                                com o dimmer da BARRA</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="1" data-bit="3" />
                                b3 >> Ignora Fotodiodo (Noite/Dia)</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="1" data-bit="2" />
                                b2 >> Apaga todos os módulos enquanto aguarda o
                                Sincronismo</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="1" data-bit="1" />
                                b1 >> Remove PWM em todos os canais da MB para usar
                                sinalizadores auxiliares</label>
                        </div>
                    </div>
                    <div class="jumptabela-section">
                        <h4 class="function-title">
                            JumpTabela #2 - Gerenciamento de Energia e Entradas
                        </h4>
                        <div class="checkbox-grid">
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="2" data-bit="7" />
                                b7 >> PWM Global de 75% (CONASP)</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="2" data-bit="6" />
                                b6 >> PWM Global de 50% (CONASP)</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="2" data-bit="5" />
                                b5 >> LOW BATTERY Management (desliga sistema com
                                Vbat&lt;11.8V)</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="2" data-bit="4" />
                                b4 >> IN3 = Entrada de acionamento rápido para luz de
                                Horn</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="2" data-bit="3" />
                                b3 >> IN4 = Entrada de acionamento dos módulos em modo
                                Background</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="2" data-bit="2" />
                                b2 >> Ativa Troca de FP pelos pinos (alguns 6
                                segundos)</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="2" data-bit="1" />
                                b1 >> Apaga TODOS os módulos enquanto aguarda o
                                Sincronismo</label>
                        </div>
                    </div>
                    <div class="jumptabela-section">
                        <h4 class="function-title">
                            JumpTabela #3 - Modo Master e Tensão
                        </h4>
                        <div class="checkbox-grid">
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="3" data-bit="7" />
                                b7 >> Ativa entradas com apenas 9 Volts</label>
                            <label class="checkbox-item"><input type="checkbox" data-jumptabela="3" data-bit="6" />
                                b6 >> MB = Master, não para / apaga pelo Sync</label>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <div class="modal-footer">
            <button
                class="btn"
                id="save-advanced-settings-btn"
                style="background-color: #27ae60">
                Salvar Definições
            </button>
        </div>
    </div>
</div>