<?php
header('Content-Type: application/json');

// --- Função para enviar erros em JSON ---
function send_json_error($message) {
    echo json_encode(['success' => false, 'error' => $message]);
    exit();
}

// --- Conexão ---
$servidor = "localhost";
$usuario = "root";
$senha = "";
$banco = "sirene_db";

$conn = new mysqli($servidor, $usuario, $senha, $banco);

if ($conn->connect_error) {
    send_json_error('Falha na conexão com o banco de dados: ' . $conn->connect_error);
}
$conn->set_charset("utf8mb4");

// --- Lógica Principal ---
$simulation_id = isset($_GET['id']) ? intval($_GET['id']) : 0;
if ($simulation_id <= 0) {
    send_json_error('ID da simulação inválido.');
}

// 1. A consulta SQL foi corrigida para selecionar as colunas que realmente existem na sua tabela.
$sql = "SELECT layout_id, block_layout, data, special_functions, jumptabela_settings FROM simulations WHERE id = ?";
$stmt = $conn->prepare($sql);

if (!$stmt) {
    send_json_error('Erro na preparação da consulta SQL: ' . $conn->error);
}

$stmt->bind_param("i", $simulation_id);
$stmt->execute();
$result = $stmt->get_result();

if ($result->num_rows > 0) {
    $row = $result->fetch_assoc();

    // 2. Monta o objeto de resposta com os dados de cada coluna.
    // O json_decode() é usado para transformar as strings do banco de volta em objetos/arrays.
    $response_data = [
        'success' => true,
        'layout_id' => $row['layout_id'], // Pega o ID do layout
        'block_layout' => json_decode($row['block_layout']),
        'data' => json_decode($row['data']),
        'special_functions' => json_decode($row['special_functions']),
        'jumptabela_settings' => json_decode($row['jumptabela_settings'])
    ];
    
    // 3. Envia o objeto JSON corretamente montado.
    echo json_encode($response_data);

} else {
    send_json_error('Nenhuma simulação encontrada com o ID fornecido: ' . $simulation_id);
}

$stmt->close();
$conn->close();
?>