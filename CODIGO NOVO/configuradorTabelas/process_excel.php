<?php
header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['error' => 'Método não permitido']);
    exit;
}

// Verifica se um arquivo foi enviado
if (!isset($_FILES['excel_file']) || $_FILES['excel_file']['error'] !== UPLOAD_ERR_OK) {
    http_response_code(400);
    echo json_encode(['error' => 'Nenhum arquivo Excel foi enviado ou ocorreu um erro no upload']);
    exit;
}

$uploadedFile = $_FILES['excel_file'];
$fileName = $uploadedFile['name'];
$tempPath = $uploadedFile['tmp_name'];

// Valida a extensão do arquivo
$allowedExtensions = ['xlsx', 'xlsm', 'xls'];
$fileExtension = strtolower(pathinfo($fileName, PATHINFO_EXTENSION));

if (!in_array($fileExtension, $allowedExtensions)) {
    http_response_code(400);
    echo json_encode(['error' => 'Tipo de arquivo não permitido. Use apenas .xlsx, .xlsm ou .xls']);
    exit;
}

try {
    // Cria um diretório temporário se não existir
    $tempDir = __DIR__ . '/temp';
    if (!is_dir($tempDir)) {
        mkdir($tempDir, 0755, true);
    }

    // Gera nomes únicos para os arquivos
    $uniqueId = uniqid();
    $excelPath = $tempDir . '/' . $uniqueId . '_' . $fileName;
    $txtPath = $tempDir . '/' . $uniqueId . '_dados_extraidos.txt';
    $jsonPath = $tempDir . '/' . $uniqueId . '_dados_extraidos.json';

    // Move o arquivo temporário para o diretório de trabalho
    if (!move_uploaded_file($tempPath, $excelPath)) {
        throw new Exception('Falha ao mover o arquivo para o diretório de processamento');
    }

    // Localiza o script Python
    $pythonScript = __DIR__ . '/../excel_parser.py';
    if (!file_exists($pythonScript)) {
        throw new Exception('Script Python excel_parser.py não encontrado em: ' . $pythonScript);
    }

    // Executa o script Python passando os caminhos como argumentos
    $command = sprintf(
        'python "%s" "%s" "%s" "%s" 2>&1',
        $pythonScript,
        $excelPath,
        $txtPath,
        $jsonPath
    );

    $output = [];
    $returnCode = 0;

    exec($command, $output, $returnCode);

    if ($returnCode !== 0) {
        throw new Exception('Erro ao executar o script Python: ' . implode("\n", $output));
    }

    // Lê o resultado JSON se foi gerado
    $result = [];
    if (file_exists($jsonPath)) {
        $jsonContent = file_get_contents($jsonPath);
        $result = json_decode($jsonContent, true);

        if ($result === null) {
            throw new Exception('Erro ao decodificar o JSON gerado pelo Python');
        }
    } else {
        throw new Exception('Arquivo JSON não foi gerado pelo script Python');
    }

    // Lê o conteúdo do arquivo TXT se disponível
    $txtContent = '';
    if (file_exists($txtPath)) {
        $txtContent = file_get_contents($txtPath);
    }

    // Limpa arquivos temporários
    @unlink($excelPath);
    @unlink($txtPath);
    @unlink($jsonPath);

    // Retorna o resultado
    echo json_encode([
        'success' => true,
        'data' => $result,
        'txt_content' => $txtContent,
        'message' => 'Arquivo Excel processado com sucesso',
        'python_output' => implode("\n", $output)
    ]);
} catch (Exception $e) {
    // Limpa arquivos temporários em caso de erro
    if (isset($excelPath)) @unlink($excelPath);
    if (isset($txtPath)) @unlink($txtPath);
    if (isset($jsonPath)) @unlink($jsonPath);

    http_response_code(500);
    echo json_encode([
        'error' => $e->getMessage(),
        'details' => isset($output) ? implode("\n", $output) : ''
    ]);
}
