<?php
// --- Coloque as MESMAS credenciais dos outros arquivos PHP aqui ---
$servidor = "localhost";
$usuario = "root";
$senha = ""; // No XAMPP padrão, a senha é vazia
$banco = "sirene_bd"; // Coloque o nome do seu banco de dados aqui
// --------------------------------------------------------------------

$conn = new mysqli($servidor, $usuario, $senha, $banco);

if ($conn->connect_error) {
    echo "<h1>FALHA NA CONEXÃO!</h1>";
    echo "<p>Erro: " . $conn->connect_error . "</p>";
    echo "<p>Verifique se o nome do banco de dados ('$banco'), o usuário e a senha estão corretos no arquivo.</p>";
} else {
    echo "<h1>SUCESSO!</h1>";
    echo "<p>A conexão com o banco de dados '$banco' foi estabelecida com sucesso.</p>";
    echo "<p>O PHP está se comunicando com o MySQL corretamente.</p>";
    $conn->close();
}
?>