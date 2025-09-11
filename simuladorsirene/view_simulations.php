<?php
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "sirene_db";
$conn = new mysqli($servername, $username, $password, $dbname);
if ($conn->connect_error) { die("Falha na conexão: " . $conn->connect_error); }
$conn->set_charset("utf8mb4");

// Adicionado 'layout_id' na consulta SQL para poder exibir o nome do modelo
$sql = "SELECT id, name, saved_at, consumo_media, consumo_pico, layout_id FROM simulations ORDER BY saved_at DESC";
$result = $conn->query($sql);
$simulations = [];
if ($result && $result->num_rows > 0) {
    while($row = $result->fetch_assoc()) { $simulations[] = $row; }
}
$conn->close();
?>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Visualizador de Simulações</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
    
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Inter', sans-serif; }
        body { background-color: #f0f2f5; color: #333; height: 100vh; overflow: hidden; }
        .header { background-color: #2c3e50; height: 40px; display: flex; align-items: center; padding: 0 24px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .logo { height: 70px; width: 100px; }
        .main-container { display: flex; height: calc(100vh - 40px); padding: 15px; gap: 15px; }
        .list-panel { width: 450px; min-width: 350px; background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); padding: 20px; display: flex; flex-direction: column; }
        .panel-title { font-size: 18px; font-weight: 600; color: #2c3e50; padding-bottom: 10px; border-bottom: 2px solid #f39c12; margin-bottom: 15px; }
        .sim-list-container { overflow-y: auto; flex-grow: 1; margin-right: -10px; padding-right: 10px; }
        .sim-list { list-style: none; }
        .search-bar-container { position: relative; margin-bottom: 15px; }
        .search-input { width: 100%; padding: 10px 15px 10px 40px; border-radius: 6px; border: 1px solid #d0d7de; font-size: 14px; transition: border-color 0.2s, box-shadow 0.2s; }
        .search-input:focus { border-color: #f39c12; box-shadow: 0 0 0 3px rgba(243, 156, 18, 0.2); outline: none; }
        .search-bar-container .fa-search { position: absolute; left: 15px; top: 50%; transform: translateY(-50%); color: #868e96; }
        .sim-card { background-color: #fff; border: 1px solid #e9ecef; border-radius: 6px; margin-bottom: 10px; padding: 15px; display: flex; flex-direction: column; gap: 10px; transition: all 0.2s ease-in-out; }
        .sim-card:hover { transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0,0,0,0.1); border-color: #f39c12; }
        .sim-card.active { border-left: 4px solid #f39c12; background-color: #fff9f0; }
        .card-main-content { display: flex; align-items: center; gap: 15px; width: 100%; }
        .card-icon { font-size: 24px; color: #f39c12; }
        .card-details { flex-grow: 1; }
        .card-details .name { font-weight: 600; font-size: 15px; color: #343a40; margin-bottom: 4px; }
        .card-details .date { font-size: 12px; color: #868e96; display: flex; align-items: center; gap: 5px; }
        .card-actions { display: flex; flex-wrap: wrap; gap: 8px; justify-content: flex-end; width: 100%; margin-top: 10px; }
        .btn { border: none; border-radius: 5px; cursor: pointer; text-decoration: none; font-size: 12px; font-weight: 500; transition: all 0.2s; padding: 6px 12px; display: flex; align-items: center; gap: 6px; }
        .btn-load { background-color: #3498db; color: white; }
        .btn-load:hover { background-color: #2980b9; }
        .btn-delete { background-color: #e74c3c; color: white; }
        .btn-delete:hover { background-color: #c0392b; }
        .btn-use-template { background-color: #27ae60; color: white; }
        .btn-use-template:hover { background-color: #229954; }
        .btn-back { background-color: #f39c12; color: #fff; width: 100%; margin-bottom: 20px; justify-content: center; }
        .btn-back:hover { background-color: #e67e22; }
        .display-panel { flex-grow: 1; background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); padding: 20px; display: flex; flex-direction: column; }
        .placeholder { display: flex; flex-direction:column; align-items: center; justify-content: center; height: 100%; color: #adb5bd; text-align: center; }
        .placeholder .fa-file-excel { font-size: 60px; margin-bottom: 20px; opacity: 0.7; }
        .placeholder-text { font-size: 18px; }
        .excel-container { display: none; flex-direction: column; height: 100%; }
        .sheet-selector-container { display: flex; flex-wrap: wrap; gap: 4px; margin-bottom: 10px; border-bottom: 2px solid #d0d7de; padding-bottom: 6px; }
        .sheet-selector-btn { background-color: #e9ecef; border: 1px solid #ced4da; padding: 5px 10px; cursor: pointer; font-size: 11px; font-weight: 500; border-radius: 4px 4px 0 0; }
        .sheet-selector-btn.active { background-color: #ffffff; color: #2c3e50; font-weight: 600; border-color: #d0d7de; border-bottom: 1px solid #fff; transform: translateY(1px); }
        .excel-grid { overflow: auto; flex-grow: 1; border: 1px solid #d0d7de; border-radius: 4px; }
        .macro-table { width: 100%; border-collapse: collapse; font-size: 9px; table-layout: fixed; }
        .macro-table th, .macro-table td { border: 1px solid #d0d7de; padding: 2px 4px; text-align: center; height: 18px; }
        .column-header { background-color: #d9e1f2; font-weight: bold; color: #333; }
        .ms25-column, .ms-column { font-weight: bold; }
        .ms25-column { color: red; background-color: #e9ecef; }
        .ms-column { color: black; background-color: #fff; }
        .cell-red { background-color: #ff0000 !important; color: white; }
        .cell-blue { background-color: #0070c0 !important; color: white; }
        .cell-white { background-color: #ffffff !important; color: black; border: 1px solid #555 !important; }
        .cell-yellow { background-color: #ffc000 !important; color: black; }
        .info-bar { display: flex; flex-wrap: wrap; gap: 15px; font-size: 11px; color: #555; margin-top: 8px; }
        .info-item { display: flex; align-items: center; gap: 5px; }
    </style>
</head>
<body>
    <div class="header">
        <img src="https://www3.facens.br/maratona/img/patrocinio-2022/Logo_Flash.png" alt="FLASH" class="logo">
    </div>
    <div class="main-container">
        <div class="list-panel">
            <a href="index.html" class="btn btn-back"><i class="fa fa-arrow-left"></i> Voltar para o Simulador</a>
            <div class="panel-title">Simulações Salvas</div>
            <div class="search-bar-container">
                <i class="fa fa-search"></i>
                <input type="text" id="search-input" class="search-input" placeholder="Pesquisar por nome...">
            </div>
            <div class="sim-list-container">
                <ul class="sim-list">
                    <?php if (!empty($simulations)): ?>
                        <?php foreach ($simulations as $sim): ?>
                            <li class="sim-card" data-name="<?php echo htmlspecialchars(strtolower($sim['name'])); ?>">
                                <div class="card-main-content">
                                    <i class="fa fa-clipboard-list card-icon"></i>
                                    <div class="card-details">
                                        <div class="name"><?php echo htmlspecialchars($sim['name']); ?></div>
                                        <div class="date"><i class="fa fa-calendar-alt"></i> Salvo em: <?php echo date('d/m/Y H:i', strtotime($sim['saved_at'])); ?></div>
                                        
                                        <div class="info-bar">
                                            <span class="info-item"><i class="fas fa-microchip" style="color: #6c757d;"></i><strong id="layout-name-<?php echo $sim['id']; ?>"></strong></span>
                                            <span class="info-item"><i class="fas fa-chart-line" style="color: #0070c0;"></i> Média: <strong><?php echo number_format($sim['consumo_media'], 3, ',', '.'); ?> A</strong></span>
                                            <span class="info-item"><i class="fas fa-bolt" style="color: #dc3545;"></i> Pico: <strong><?php echo number_format($sim['consumo_pico'], 3, ',', '.'); ?> A</strong></span>
                                        </div>
                                        
                                        <script>
                                            document.addEventListener('DOMContentLoaded', () => {
                                                const layouts = [ // Este array mapeia o ID do layout para o nome amigável
                                                    { id: 's_slim_24_default', name: 'S-SLIM 24m Padrão'}, { id: 'primus40p12mCol', name: 'Primus 40p 12m Col'}, 
                                                    { id: 'primus40p12mRef', name: 'Primus 40p 12m Ref'}, { id: 'primus40p14mCol', name: 'Primus 40p 14m Col'}, 
                                                    { id: 'primus40p14mRef', name: 'Primus 40p 14m Ref'}, { id: 'primus40p16mCol', name: 'Primus 40p 16m Col'}, 
                                                    { id: 'primus40p16mRef', name: 'Primus 40p 16m Ref'}, { id: 'primus40p18mCol', name: 'Primus 40p 18m Col'}, 
                                                    { id: 'primus40p18mRef', name: 'Primus 40p 18m Ref'}, { id: 'primus40p22mRef', name: 'Primus 40p 22m Ref'}, 
                                                    { id: 'primus54p10mCol', name: 'Primus 54p 10m Col'}, { id: 'primus54p12mRef', name: 'Primus 54p 12m Ref'}, 
                                                    { id: 'primus54p14mCol', name: 'Primus 54p 14m Col'}, { id: 'primus54p14mRef', name: 'Primus 54p 14m Ref'}, 
                                                    { id: 'primus54p16mCol', name: 'Primus 54p 16m Col'}, { id: 'primus54p18mRef', name: 'Primus 54p 18m Ref'}, 
                                                    { id: 'primus54p20mRef', name: 'Primus 54p 20m Ref'}, { id: 'primus54p22mCol', name: 'Primus 54p 22m Col'}, 
                                                    { id: 'primus54p22mRef', name: 'Primus 54p 22m Ref'}, { id: 'primus54p26mRef', name: 'Primus 54p 26m Ref'}, 
                                                    { id: 'primus61p13mCol', name: 'Primus 61p 13m Col'}, { id: 'primus61p13mRef', name: 'Primus 61p 13m Ref'}, 
                                                    { id: 'primus61p15mRef', name: 'Primus 61p 15m Ref'}, { id: 'primus61p18mCol', name: 'Primus 61p 18m Col'}, 
                                                    { id: 'primus61p18mRef', name: 'Primus 61p 18m Ref'}, { id: 'primus61p20mRef', name: 'Primus 61p 20m Ref'}, 
                                                    { id: 'primus61p22mRef', name: 'Primus 61p 22m Ref'}, { id: 'primus61p24mCol', name: 'Primus 61p 24m Col'}, 
                                                    { id: 'primus61p24mRef', name: 'Primus 61p 24m Ref'}, { id: 'primus61p28mRef', name: 'Primus 61p 28m Ref'}, 
                                                    { id: 'venus11m', name: 'Venus 11m'}, { id: 'venus14m', name: 'venus 14m'}, { id: 'venus15m', name: 'venus 15m'}, 
                                                    { id: 'venus16m', name: 'venus 16m'}, { id: 'venus19m', name: 'venus 19m'}, { id: 'venus20m', name: 'venus 20m'}, 
                                                    { id: 'venus22m', name: 'venus 22m'},
                                                ];
                                                const layoutId = "<?php echo $sim['layout_id']; ?>";
                                                const layoutNameElement = document.getElementById('layout-name-<?php echo $sim['id']; ?>');
                                                const foundLayout = layouts.find(l => l.id === layoutId);
                                                if (layoutNameElement) {
                                                    layoutNameElement.textContent = foundLayout ? foundLayout.name : 'Modelo Padrão';
                                                }
                                            });
                                        </script>
                                    </div>
                                </div>
                                <div class="card-actions">
                                    <button class="btn btn-use-template" data-id="<?php echo $sim['id']; ?>">
                                        <i class="fa fa-copy"></i> Usar como Template
                                    </button>
                                    <button class="btn btn-load" data-id="<?php echo $sim['id']; ?>">
                                        <i class="fa fa-eye"></i> Visualizar
                                    </button>
                                    <button class="btn btn-delete" data-id="<?php echo $sim['id']; ?>">
                                        <i class="fa fa-trash"></i> Deletar
                                    </button>
                                </div>
                            </li>
                        <?php endforeach; ?>
                    <?php else: ?>
                        <li style="text-align:center; padding: 20px; color: #868e96;">Nenhuma simulação salva.</li>
                    <?php endif; ?>
                </ul>
            </div>
        </div>
        <div class="display-panel">
            <div id="placeholder" class="placeholder">
                <i class="far fa-file-excel"></i>
                <span class="placeholder-text">Selecione uma simulação na lista para visualizar.</span>
            </div>
            <div id="excel-container" class="excel-container">
                <div id="sheet-selector-container" class="sheet-selector-container"></div>
                <div class="excel-grid">
                    <table class="macro-table">
                        <thead id="macro-header"></thead>
                        <tbody id="macro-body"></tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script>
    document.addEventListener('DOMContentLoaded', () => {
        const placeholder = document.getElementById('placeholder');
        const excelContainer = document.getElementById('excel-container');
        const sheetSelectorContainer = document.getElementById('sheet-selector-container');
        const macroHeader = document.getElementById('macro-header');
        const macroBody = document.getElementById('macro-body');
        const searchInput = document.getElementById('search-input');
        
        let fullSimulationData = null;

        document.querySelectorAll('.btn-load').forEach(button => {
            button.addEventListener('click', (e) => {
                const simId = e.currentTarget.dataset.id;
                document.querySelectorAll('.sim-card').forEach(card => card.classList.remove('active'));
                e.currentTarget.closest('.sim-card').classList.add('active');
                loadSimulationData(simId, false); // false = não é para template
            });
        });

        document.querySelectorAll('.btn-use-template').forEach(button => {
            button.addEventListener('click', (e) => {
                const simId = e.currentTarget.dataset.id;
                loadSimulationData(simId, true); // true = é para template
            });
        });

        searchInput.addEventListener('input', (e) => {
            const searchTerm = e.target.value.toLowerCase();
            document.querySelectorAll('.sim-card').forEach(card => {
                const cardName = card.dataset.name;
                card.style.display = cardName.includes(searchTerm) ? 'flex' : 'none';
            });
        });

        async function loadSimulationData(id, isForTemplate) {
            const actionButton = document.querySelector(`.btn[data-id='${id}']${isForTemplate ? '.btn-use-template' : '.btn-load'}`);
            const originalText = actionButton.innerHTML;
            actionButton.innerHTML = '<i class="fa fa-spinner fa-spin"></i>';
            actionButton.disabled = true;
            
            if (!isForTemplate) {
                placeholder.style.display = 'flex';
                excelContainer.style.display = 'none';
                placeholder.innerHTML = `<i class="fa fa-spinner fa-spin"></i><span class="placeholder-text" style="margin-left:10px;">Carregando...</span>`;
            }

            try {
                const response = await fetch(`get_simulation_data.php?id=${id}`);
                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                const result = await response.json();

                if (!result.success) throw new Error(result.error);
                
                if (isForTemplate) {
                    // 'result' já tem a estrutura correta e completa para o template
                    localStorage.setItem('simulation_template', JSON.stringify(result));
                    window.location.href = 'index.html';
                } else {
                    // Os dados da planilha estão agora diretamente em 'result.data'
                    fullSimulationData = result.data; 
                    placeholder.style.display = 'none';
                    excelContainer.style.display = 'flex';
                    renderSimulationDisplay(fullSimulationData);
                }

            } catch (error) {
                console.error('Falha ao carregar dados:', error);
                if (!isForTemplate) {
                    placeholder.innerHTML = `<i class="fa fa-exclamation-triangle"></i><span class="placeholder-text">Falha ao carregar.</span>`;
                } else {
                    alert('Não foi possível carregar o template. Tente novamente.');
                }
            } finally {
                 if(actionButton){
                    actionButton.innerHTML = originalText;
                    actionButton.disabled = false;
                 }
            }
        }
        
        document.querySelectorAll('.btn-delete').forEach(button => {
            button.addEventListener('click', (e) => {
                e.stopPropagation(); 
                const card = e.currentTarget.closest('.sim-card');
                const simId = e.currentTarget.dataset.id;
                const simName = card.querySelector('.name').textContent;

                if (confirm(`Tem certeza que deseja deletar a simulação "${simName}"? Esta ação não pode ser desfeita.`)) {
                    deleteSimulation(simId, card);
                }
            });
        });

        async function deleteSimulation(id, cardElement) {
            const formData = new FormData();
            formData.append('id', id);

            try {
                const response = await fetch('delete_simulation.php', {
                    method: 'POST',
                    body: formData
                });
                const result = await response.json();

                if (result.success) {
                    cardElement.style.transition = 'opacity 0.3s ease';
                    cardElement.style.opacity = '0';
                    setTimeout(() => {
                        cardElement.remove();
                        if(cardElement.classList.contains('active')) {
                            placeholder.innerHTML = `<i class="far fa-file-excel"></i><span class="placeholder-text">Selecione uma simulação na lista para visualizar.</span>`;
                            placeholder.style.display = 'flex';
                            excelContainer.style.display = 'none';
                        }
                    }, 300);
                } else {
                    alert('Erro ao deletar: ' + result.error);
                }
            } catch(error) {
                console.error("Erro na requisição de deleção:", error);
                alert("Ocorreu um erro de rede ao tentar deletar a simulação.");
            }
        }

        function renderSimulationDisplay(data) {
            const sheetNames = Object.keys(data);
            createSheetTabs(sheetNames);
            if (sheetNames.length > 0) renderSheetContent(sheetNames[0]);
        }
        function createSheetTabs(sheetNames) {
            sheetSelectorContainer.innerHTML = '';
            sheetNames.forEach(name => {
                const button = document.createElement('button');
                button.className = 'sheet-selector-btn';
                button.textContent = name;
                button.addEventListener('click', () => {
                    document.querySelectorAll('.sheet-selector-btn').forEach(btn => btn.classList.remove('active'));
                    button.classList.add('active');
                    renderSheetContent(name);
                });
                sheetSelectorContainer.appendChild(button);
            });
            if (sheetSelectorContainer.firstChild) sheetSelectorContainer.firstChild.classList.add('active');
        }
        function renderSheetContent(sheetName) {
            const sheetData = fullSimulationData[sheetName];
            if (!sheetData) return;
            macroHeader.innerHTML = '';
            const trHead = document.createElement('tr');
            trHead.innerHTML = `<th class="column-header">#</th><th class="column-header">ms/25</th><th class="column-header">ms</th>`;
            for (let i = 1; i <= 24; i++) trHead.innerHTML += `<th class="column-header">${i}</th>`;
            macroHeader.appendChild(trHead);
            macroBody.innerHTML = '';
            sheetData.forEach(rowData => {
                const tr = document.createElement('tr');
                const msValue = parseInt(rowData.ms, 10);
                const ms25Value = isNaN(msValue) ? '' : (msValue === 0 ? '0' : Math.round(msValue / 25));

                tr.innerHTML = `<td>${rowData.row || ''}</td><td class="ms25-column">${ms25Value}</td><td class="ms-column">${rowData.ms || ''}</td>`;
                const values = rowData.values || [];
                for (let i = 0; i < 24; i++) {
                    const value = values[i] || '';
                    const td = document.createElement('td');
                    td.textContent = value;
                    updateCellColor(td);
                    tr.appendChild(td);
                }
                macroBody.appendChild(tr);
            });
        }
        function updateCellColor(cell) {
            const value = cell.textContent.trim();
            cell.className = '';
            if (value === '1') cell.classList.add('cell-red');
            else if (value === '2') cell.classList.add('cell-blue');
            else if (value === '3') cell.classList.add('cell-white');
            else if (value === '4') cell.classList.add('cell-yellow');
        }
    });
    </script>
</body>
</html>