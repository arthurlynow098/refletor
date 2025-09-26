<!DOCTYPE html>
<html lang="pt-BR">

<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Simulação Sirene - FLASH (Versão Final)</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet" />
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css" />

  <!-- Bibliotecas Externas -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/gif.js/0.2.0/gif.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

  <!-- Estilos da Aplicação -->
  <style>
    /* O ideal seria ter um arquivo .css separado, mas manteremos o include por enquanto */
    <?php include '../css/styles.css'; ?>
  </style>
</head>

<body>
  <ul id="color-context-menu">
    <li data-color="red">Vermelha</li>
    <li data-color="green">Verde</li>
    <li data-color="blue">Azul</li>
    <li data-color="white">Branca</li>
    <li data-color="yellow">Âmbar</li>
    <hr />
    <li data-color="clear">Limpar Cor</li>
  </ul>

  <div id="layout-panel-modal" class="modal-overlay">
    <div class="modal-content" style="width: 80%; max-width: 1200px">
      <div class="modal-header">
        <div class="panel-title" style="margin-bottom: 0; border-bottom: none">
          <i class="fa fa-th-large"></i> Selecione um Modelo de Sinalizador
        </div>
        <button class="close-button" id="close-layout-panel-btn">
          &times;
        </button>
      </div>
      <div class="modal-body" id="layout-panel-body"></div>
    </div>
  </div>

  <?php include 'definicoes.php'; ?>

  <div id="config-modal" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <div class="panel-title" style="margin-bottom: 0; border-bottom: none">
          Dados de Configuração
        </div>
        <button class="close-button" id="close-modal-btn">&times;</button>
      </div>
      <div class="modal-body">
        <div class="form-group">
          <label class="form-label">Configuração de Barra</label><input
            type="text"
            class="config-input"
            id="config-barra"
            value="FLASH STD" />
        </div>
        <div class="form-group">
          <label class="form-label">Modelo da Barra</label><input
            type="text"
            class="config-input"
            id="modelo-barra"
            value="Linha S-Slim" />
        </div>
        <div class="form-group">
          <label class="form-label">Sub-família</label><input
            type="text"
            class="config-input"
            id="sub-familia"
            value="S-SLIM Bicolor" />
        </div>
        <div class="form-group">
          <label class="form-label">Tamanho da Barra</label><input
            type="text"
            class="config-input"
            id="tamanho-barra"
            value="1200mm" />
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn" id="save-and-close-modal-btn">
          Salvar e Fechar
        </button>
      </div>
    </div>
  </div>


  <?php include '../navbar.php'; ?>
  <div class="titulo">
    <div class="container">
      <h1>CONFIGURADOR DE TABELAS</h1>
    </div>
  </div>

  <div class="main-container">
    <?php include 'mainLeft.php'; ?>
    <?php include 'mainRight.php'; ?>
  </div>

  <!-- Scripts da Aplicação -->
  <script>
    <?php
    include '../js/mapaBarras.js';
    include '../js/declaracoes.js';
    include '../js/funcoes.js';
    ?>
  </script>

</body>

</html>