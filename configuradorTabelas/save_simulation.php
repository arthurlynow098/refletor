<?php
// save_simulation.php (VERSÃO CORRIGIDA PARA SALVAR O LAYOUT)

header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-cache, must-revalidate');

$servername = "localhost";
$username = "root";
$password = "";
$dbname = "sirene_db";

$response = [];

try {
    $inputJSON = file_get_contents('php://input');
    if ($inputJSON === false) {
        throw new Exception("Não foi possível ler os dados da requisição.");
    }

    $data = json_decode($inputJSON, true);
    if (json_last_error() !== JSON_ERROR_NONE) {
        throw new Exception("Erro ao decodificar o JSON: " . json_last_error_msg());
    }

    // ADICIONADO: Validação dos novos campos de layout
    $required_fields = ['name', 'data', 'consumoMedia', 'consumoPico', 'special_functions', 'jumptabela_settings', 'layout_id', 'block_layout'];
    foreach ($required_fields as $field) {
        if (!isset($data[$field])) {
            throw new Exception("O campo obrigatório '$field' não foi recebido.");
        }
    }

    $dsn = "mysql:host=$servername;dbname=$dbname;charset=utf8mb4";
    $options = [
        PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        PDO::ATTR_EMULATE_PREPARES   => false,
    ];
    $pdo = new PDO($dsn, $username, $password, $options);

    // ALTERADO: Query de inserção com os novos campos
    $sql = "INSERT INTO simulations (name, layout_id, block_layout, data, consumo_media, consumo_pico, special_functions, jumptabela_settings) 
            VALUES (:name, :layout_id, :block_layout, :data, :consumo_media, :consumo_pico, :special_functions, :jumptabela_settings)";
    
    $stmt = $pdo->prepare($sql);

    // ALTERADO: Associa os novos parâmetros
    $stmt->execute([
        ':name' => $data['name'],
        ':layout_id' => $data['layout_id'], // NOVO
        ':block_layout' => json_encode($data['block_layout']), // NOVO
        ':data' => json_encode($data['data']),
        ':consumo_media' => $data['consumoMedia'],
        ':consumo_pico' => $data['consumoPico'],
        ':special_functions' => json_encode($data['special_functions']),
        ':jumptabela_settings' => json_encode($data['jumptabela_settings'])
    ]);

    $response = ['success' => true, 'message' => 'Simulação salva com sucesso!'];

} catch (Exception $e) {
    http_response_code(500);
    $response = ['success' => false, 'error' => $e->getMessage()];
}

echo json_encode($response);
?>