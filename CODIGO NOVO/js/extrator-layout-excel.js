document.addEventListener("DOMContentLoaded", () => {
  const inputArquivo = document.getElementById("arquivoExcel");
  const botaoExtrair = document.getElementById("botaoExtrair");
  const areaResultado = document.getElementById("resultado");

  botaoExtrair.addEventListener("click", async () => {
    const arquivo = inputArquivo.files[0];

    if (!arquivo) {
      areaResultado.textContent = "Por favor, selecione um arquivo primeiro.";
      return;
    }

    botaoExtrair.disabled = true;
    botaoExtrair.textContent = "Processando...";
    areaResultado.textContent = "Lendo e processando o arquivo...";

    try {
      const resultado = await processExcelLayoutJS(arquivo);
      // Exibe o resultado formatado como JSON
      areaResultado.textContent = JSON.stringify(resultado, null, 2);
    } catch (error) {
      console.error("Erro ao processar o arquivo Excel:", error);
      areaResultado.textContent = `Ocorreu um erro: ${error.message}`;
    } finally {
      botaoExtrair.disabled = false;
      botaoExtrair.textContent = "Extrair Layout";
    }
  });

  /**
   * Lê um arquivo Excel e extrai o layout completo, replicando a lógica do script Python.
   * @param {File} file O arquivo Excel selecionado pelo usuário.
   * @returns {Promise<object>} Uma promessa que resolve com o objeto de layout extraído.
   */
  async function processExcelLayoutJS(file) {
    const COLOR_MAP = {
      VM: "vermelho",
      AZ: "azul",
      BR: "branco",
      AB: "ambar",
    };

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error("Erro ao ler o arquivo."));

      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const sheetNames = workbook.SheetNames;

          // Procura pela planilha adequada (prioriza "POL")
          let targetSheetName =
            sheetNames.find((name) => name.toUpperCase().includes("POL")) ||
            sheetNames[0];
          if (!targetSheetName)
            return reject(new Error("Nenhuma planilha encontrada."));

          const worksheet = workbook.Sheets[targetSheetName];
          if (!worksheet)
            return reject(
              new Error(`A planilha '${targetSheetName}' não foi encontrada.`)
            );

          // --- INÍCIO DA LÓGICA DE EXTRAÇÃO COMPLEXA ---

          // 1. Extrair informações do produto (SIN VISUAL)
          let productInfo = {
            codigo_firmware: "",
            produto: "",
            modelo: "",
            cores_leds: "()",
            funcoes_extras: "",
            cores_tampas: "",
            tamanho_polegadas: "",
            quantidade_modulos: "0",
            numero_leds: "0",
            cores_leds_lista: [],
          };

          const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:Z100");
          let infoEncontrada = false;

          for (let r = range.s.r; r <= range.e.r; ++r) {
            for (let c = range.s.c; c <= range.e.c; ++c) {
              const cellAddress = XLSX.utils.encode_cell({ r, c });
              const cell = worksheet[cellAddress];
              if (cell && cell.v && String(cell.v).includes("SIN VISUAL")) {
                const info = String(cell.v);
                infoEncontrada = true;
                const primusMatch = info.match(
                  /([\d\w\.]+)\s*-\s*SIN VISUAL\s+(PRIMUS|ARES)\s+(\w+)\s+(\d+)P\s+(\d+)M\s+(\d+)L\s+(\([^)]+\))\s+(\w+)\s+TP\s+([\w/]+)/i
                );
                if (primusMatch) {
                  productInfo.codigo_firmware = primusMatch[1];
                  productInfo.produto = primusMatch[2];
                  productInfo.modelo = primusMatch[3];
                  productInfo.tamanho_polegadas = primusMatch[4];
                  productInfo.quantidade_modulos = primusMatch[5];
                  productInfo.numero_leds = primusMatch[6];
                  productInfo.cores_leds = primusMatch[7];
                  productInfo.funcoes_extras = primusMatch[8];
                  productInfo.cores_tampas = primusMatch[9];
                  const cores_str = primusMatch[7].replace(/[()]/g, "");
                  productInfo.cores_leds_lista = cores_str
                    .split("+")
                    .map((c) => c.trim())
                    .filter(Boolean);
                }
                break;
              }
            }
            if (infoEncontrada) break;
          }

          // 2. Mapear cores e canais dos módulos
          const posicoes_modulos = [
            "O4:P5",
            "Q4:R5",
            "S4:T5",
            "V4:W5",
            "X4:Y5",
            "Z4:AA5",
            "AB6:AC7",
            "AD10:AE11",
            "AD14:AE15",
            "AD18:AE19",
            "AB22:AC23",
            "Z24:AA25",
            "X24:Y25",
            "V24:W25",
            "S24:T25",
            "Q24:R25",
            "O24:P25",
            "L24:M25",
            "J24:K25",
            "H24:I25",
            "F22:G23",
            "D18:E19",
            "D14:E15",
            "D10:E11",
            "F6:G7",
            "H4:I5",
            "J4:K5",
            "L4:M5",
          ];

          const POSICOES_CANAIS = {
            "O4:P5": [2, 0],
            "Q4:R5": [2, 0],
            "S4:T5": [2, 0],
            "V4:W5": [2, 0],
            "X4:Y5": [2, 0],
            "Z4:AA5": [2, 0],
            "AB6:AC7": [2, 0],
            "AD10:AE11": [2, 0],
            "AD14:AE15": [0, -2],
            "AD18:AE19": [-2, 0],
            "AB22:AC23": [-2, 0],
            "Z24:AA25": [-2, 0],
            "X24:Y25": [-2, 0],
            "V24:W25": [-2, 0],
            "S24:T25": [-2, 0],
            "Q24:R25": [-2, 0],
            "O24:P25": [-2, 0],
            "L24:M25": [-2, 0],
            "J24:K25": [-2, 0],
            "H24:I25": [-2, 0],
            "F22:G23": [-2, 0],
            "D18:E19": [-2, 0],
            "D14:E15": [0, 2],
            "D10:E11": [2, 0],
            "F6:G7": [2, 0],
            "H4:I5": [2, 0],
            "J4:K5": [2, 0],
            "L4:M5": [2, 0],
          };

          const mapeamento_cores_modulos = [];

          posicoes_modulos.forEach((range_str, index) => {
            const modulo_num = index + 1;
            const mapeamento = [];

            try {
              const range = XLSX.utils.decode_range(range_str);
              const canal_info = POSICOES_CANAIS[range_str];

              for (let r = range.s.r; r <= range.e.r; r++) {
                for (let c = range.s.c; c <= range.e.c; c++) {
                  const cellAddress = XLSX.utils.encode_cell({ r, c });
                  const cell = worksheet[cellAddress];

                  if (cell && cell.v) {
                    const texto_celula = String(cell.v).trim().toUpperCase();

                    for (const [codigo, nome_cor] of Object.entries(
                      COLOR_MAP
                    )) {
                      if (texto_celula.includes(codigo)) {
                        let coord_canal_final = null;
                        let numero_canal = null;

                        if (canal_info) {
                          const [row_offset, col_offset] = canal_info;
                          const nova_row = r + row_offset;
                          const nova_col = c + col_offset;

                          if (nova_row >= 0 && nova_col >= 0) {
                            coord_canal_final = XLSX.utils.encode_cell({
                              r: nova_row,
                              c: nova_col,
                            });
                            const canal_cell = worksheet[coord_canal_final];
                            if (
                              canal_cell &&
                              canal_cell.v &&
                              !isNaN(canal_cell.v)
                            ) {
                              numero_canal = parseInt(canal_cell.v, 10);
                            }
                          }
                        }

                        mapeamento.push({
                          cor: nome_cor,
                          canal: numero_canal,
                          canal_encontrado_em: coord_canal_final,
                        });
                      }
                    }
                  }
                }
              }
            } catch (e) {
              console.warn(
                `Erro ao processar módulo ${modulo_num} (${range_str}):`,
                e
              );
            }

            mapeamento_cores_modulos.push({
              modulo: modulo_num,
              posicao: range_str,
              mapeamento: mapeamento,
            });
          });

          const modulos_ativos = mapeamento_cores_modulos.filter(
            (mod) =>
              mod.mapeamento &&
              mod.mapeamento.length > 0 &&
              mod.mapeamento.some(
                (map) => map.canal !== null && map.canal !== 0
              )
          ).length;

          // 3. Montar o resultado final
          const resultado_final = {
            ...productInfo,
            mapeamento_cores_modulos: mapeamento_cores_modulos,
            quantidade_modulos_original: productInfo.quantidade_modulos,
            quantidade_modulos_ativos: modulos_ativos,
          };

          // Resolve a promessa com os dados completos
          resolve(resultado_final);
        } catch (error) {
          console.error("Erro no processamento JavaScript:", error);
          reject(error);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }
});
