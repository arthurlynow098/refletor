<?php
// delete_simulation.php (VERSÃO CORRIGIDA E MAIS SEGURA)

header('Content-Type: application/json; charset=utf-8');

$servername = "localhost";
$username = "root";
$password = "";
$dbname = "sirene_db";

if (!isset($_POST['id']) || !filter_var($_POST['id'], FILTER_VALIDATE_INT)) {
    http_response_code(400);
    echo json_encode(['success' => false, 'error' => 'ID da simulação inválido ou não fornecido.']);
    exit;
}
$id = (int)$_POST['id'];

try {
    $dsn = "mysql:host=$servername;dbname=$dbname;charset=utf8mb4";
    $options = [ PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION ];
    $pdo = new PDO($dsn, $username, $password, $options);

    $stmt = $pdo->prepare("DELETE FROM simulations WHERE id = ?");
    $stmt->execute([$id]);

    if ($stmt->rowCount() > 0) {
        echo json_encode(['success' => true, 'message' => 'Simulação deletada com sucesso!']);
    } else {
        echo json_encode(['success' => false, 'error' => 'Nenhuma simulação encontrada com o ID fornecido.']);
    }

} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode(['success' => false, 'error' => 'Erro no servidor: ' . $e->getMessage()]);
}
?>