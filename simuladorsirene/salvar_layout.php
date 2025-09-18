<?php
header('Content-Type: application/json');

$servidor = "localhost";
$usuario = "root";
$senha = ""; 
$banco = "sirene_db";

$conn = new mysqli($servidor, $usuario, $senha, $banco);

if ($conn->connect_error) {
    echo json_encode(['success' => false, 'error' => 'Falha na conexão: ' . $conn->connect_error]);
    exit();
}

$dados = json_decode(file_get_contents('php://input'), true);

if (!$dados || !isset($dados['layout_id'])) {
    echo json_encode(['success' => false, 'error' => 'Dados inválidos ou ID do layout ausente.']);
    exit();
}

$layout_id = $dados['layout_id'];
$nome_layout = $dados['nome']; // Vamos precisar do nome e imagem para o caso de ser a primeira vez que salvamos
$imagem_base = $dados['imagem'];
$dados_layout = json_encode($dados['layout_data'] ?? []);

// Este comando mágico faz tudo:
// 1. Tenta INSERIR um novo layout com os dados fornecidos.
// 2. Se um layout com o mesmo `layout_id` (nossa chave primária) JÁ EXISTIR,
//    ele executa a parte UPDATE, atualizando apenas os dados dos blocos.
$stmt = $conn->prepare("
    INSERT INTO layouts (layout_id, nome_layout, imagem_base, dados_layout) 
    VALUES (?, ?, ?, ?)
    ON DUPLICATE KEY UPDATE dados_layout = VALUES(dados_layout)
");

$stmt->bind_param("ssss", $layout_id, $nome_layout, $imagem_base, $dados_layout);

if ($stmt->execute()) {
    echo json_encode(['success' => true, 'message' => 'Layout salvo com sucesso!']);
} else {
    echo json_encode(['success' => false, 'error' => 'Erro ao salvar o layout: ' . $stmt->error]);
}

$stmt->close();
$conn->close();
?>