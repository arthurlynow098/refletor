<?php
// get_simulation_data.php (VERSÃO FINAL CORRIGIDA)

// Define o cabeçalho da resposta como JSON e evita cache
header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-cache, must-revalidate');

// --- Detalhes da conexão ---
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "sirene_db";
// -------------------------

// Valida o ID recebido
if (!isset($_GET['id']) || !filter_var($_GET['id'], FILTER_VALIDATE_INT)) {
    http_response_code(400); // Bad Request
    echo json_encode(['success' => false, 'error' => 'ID da simulação inválido ou não fornecido.']);
    exit;
}
$id = (int)$_GET['id'];

try {
    // Conexão com o banco de dados usando PDO
    $dsn = "mysql:host=$servername;dbname=$dbname;charset=utf8mb4";
    $options = [
        PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        PDO::ATTR_EMULATE_PREPARES   => false,
    ];
    $pdo = new PDO($dsn, $username, $password, $options);

    // Seleciona TODAS as colunas necessárias para restaurar o estado completo da simulação.
    $stmt = $pdo->prepare("SELECT `layout_id`, `block_layout`, `data`, `special_functions`, `jumptabela_settings` FROM simulations WHERE id = ?");
    $stmt->execute([$id]);
    $simulation = $stmt->fetch();

    if ($simulation) {
        // Monta a resposta JSON "plana" (sem aninhar) que o JavaScript do index.html espera
        $response = [
            'success'               => true,
            'layout_id'             => $simulation['layout_id'],
            'block_layout'          => json_decode($simulation['block_layout']),
            'data'                  => json_decode($simulation['data']),
            'special_functions'     => json_decode($simulation['special_functions']),
            'jumptabela_settings'   => json_decode($simulation['jumptabela_settings'])
        ];
        echo json_encode($response);
    } else {
        http_response_code(404); // Not Found
        echo json_encode(['success' => false, 'error' => 'Simulação não encontrada.']);
    }

} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode(['success' => false, 'error' => 'Erro no servidor: ' . $e->getMessage()]);
}
?>