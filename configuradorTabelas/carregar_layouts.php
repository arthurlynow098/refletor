<?php
header('Content-Type: application/json');

// --- Conexão com o Banco de Dados (ajuste com suas credenciais) ---
$servidor = "localhost";
$usuario = "root";
$senha = ""; // No XAMPP padrão, a senha é vazia
$banco = "sirene_db"; // <<-- COLOQUE O MESMO NOME DO BANCO AQUI

$conn = new mysqli($servidor, $usuario, $senha, $banco);

if ($conn->connect_error) {
    echo json_encode([]); // Retorna um array vazio em caso de erro
    exit();
}
// --------------------------------------------------------------------

$layouts = [];
$resultado = $conn->query("SELECT layout_id, nome_layout, imagem_base, categoria, dados_layout FROM layouts");

if ($resultado) {
    while ($linha = $resultado->fetch_assoc()) {
        // Decodifica a string JSON para um objeto/array PHP antes de enviar
        // Isso é crucial para o JavaScript receber os dados no formato correto
        $linha['dados_layout'] = json_decode($linha['dados_layout']);
        $layouts[] = $linha;
    }
}

echo json_encode($layouts);

$conn->close();
?>