"use server";

import * as XLSX from "xlsx";

export const processarExcel = async (
  file: File,
  abaOrigem: string,
  colunas: string[],
  abaDestino: string,
) => {
  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
  });

  // ============================================================
  // ABA DE ORIGEM
  // ============================================================

  const worksheetOrigem = workbook.Sheets[abaOrigem];

  if (!worksheetOrigem) {
    throw new Error(
      `A aba "${abaOrigem}" não foi encontrada.`,
    );
  }

  // ============================================================
  // DADOS DA ORIGEM
  // ============================================================

  const dadosOrigem =
    XLSX.utils.sheet_to_json<unknown[]>(
      worksheetOrigem,
      {
        header: 1,
        defval: "",
      },
    );

  // ============================================================
  // CONVERTE AS COLUNAS
  // A, B, C -> 0, 1, 2
  // ============================================================

  const indices = colunas.map((coluna) =>
    XLSX.utils.decode_col(
      coluna.trim().toUpperCase(),
    ),
  );

  if (indices.length === 0) {
    throw new Error(
      "Nenhuma coluna foi informada.",
    );
  }

  // ============================================================
  // SELECIONA AS COLUNAS DA ORIGEM
  // ============================================================

  const dadosSelecionados = dadosOrigem.map(
    (linha) =>
      indices.map(
        (indice) => linha[indice] ?? "",
      ),
  );

  // ============================================================
  // ABA DE DESTINO
  // ============================================================

  let worksheetDestino =
    workbook.Sheets[abaDestino];

  if (!worksheetDestino) {
    worksheetDestino =
      XLSX.utils.aoa_to_sheet([]);

    workbook.Sheets[abaDestino] =
      worksheetDestino;

    workbook.SheetNames.push(abaDestino);
  }

  // ============================================================
  // DADOS EXISTENTES NA DESTINO
  // ============================================================

  const dadosDestino =
    XLSX.utils.sheet_to_json<unknown[]>(
      worksheetDestino,
      {
        header: 1,
        defval: "",
      },
    );

  // ============================================================
  // CRIA UM SET COM AS LINHAS EXISTENTES
  //
  // Exemplo:
  //
  // ["João", "SP"] -> "João|SP"
  // ============================================================

  const linhasExistentes = new Set(
    dadosDestino.map((linha) =>
      linha
        .map((valor) =>
          String(valor).trim(),
        )
        .join("|"),
    ),
  );

  // ============================================================
  // FILTRA SOMENTE O QUE NÃO EXISTE
  // ============================================================

  const novasLinhas =
    dadosSelecionados.filter((linha) => {
      const chave = linha
        .map((valor) =>
          String(valor).trim(),
        )
        .join("|");

      if (linhasExistentes.has(chave)) {
        return false;
      }

      // Adiciona ao Set para evitar
      // duplicados dentro da própria origem
      linhasExistentes.add(chave);

      return true;
    });

  // ============================================================
  // ADICIONA SOMENTE AS LINHAS NOVAS
  // ============================================================

  if (novasLinhas.length > 0) {
    XLSX.utils.sheet_add_aoa(
      worksheetDestino,
      novasLinhas,
      {
        origin: -1,
      },
    );
  }

  // ============================================================
  // GERA O ARQUIVO
  // ============================================================

  const nomeOriginal =
    file.name.replace(
      /\.[^/.]+$/,
      "",
    );

  const output = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "buffer",
  });

  return {
    nome: `${nomeOriginal}_mesclada.xlsx`,
    arquivo: output.toString("base64"),
  };
};