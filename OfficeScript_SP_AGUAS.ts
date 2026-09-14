/**
 * SP ÁGUAS — ponte Power Automate -> Excel.
 *
 * O fluxo Power Automate chama este script com:
 *   action = GET_SNAPSHOT | REPLACE_SNAPSHOT
 *   payloadJson = JSON completo recebido pelo fluxo.
 *
 * O script não usa Graph e não precisa de credenciais.
 */

const COLUNAS: Record<string, string[]> = {
  usuarios: ["id", "nome", "usuario", "email", "senha_hash", "perfil", "ativo", "aprovado", "criado_em"],
  demandas: ["id", "numero_processo", "numero_etc", "origem", "assunto", "area", "responsavel", "data_recebimento", "prazo_area", "prazo_fatal", "situacao", "prioridade", "observacoes", "doe_data", "doe_edicao", "doe_secao", "doe_palavra_chave", "doe_publicacao", "doe_url", "criado_por", "criado_em", "atualizado_em"],
  historico: ["id", "demanda_id", "usuario_id", "acao", "descricao", "data_hora"]
};

function main(workbook: ExcelScript.Workbook, action: string, payloadJson: string): string {
  try {
    const request = JSON.parse(payloadJson || "{}");

    if (action === "GET_SNAPSHOT") {
      return JSON.stringify({ status: "OK", action, snapshot: lerSnapshot(workbook), version: obterVersao(workbook) });
    }

    if (action === "REPLACE_SNAPSHOT") {
      const snapshot = request.snapshot || {};
      const expectedVersion = String(request.expectedVersion || "");
      const atualVersion = obterVersao(workbook);
      if (expectedVersion && atualVersion && expectedVersion !== atualVersion) {
        return JSON.stringify({ status: "ERRO", httpStatus: 412, mensagem: "O Excel foi alterado depois da leitura do SP ÁGUAS.", version: atualVersion });
      }
      validarSnapshot(snapshot);
      const novaVersao = new Date().toISOString();
      gravarSnapshot(workbook, snapshot);
      definirVersao(workbook, novaVersao);
      return JSON.stringify({ status: "OK", action, version: novaVersao });
    }

    throw new Error(`Ação não suportada: ${action}`);
  } catch (erro) {
    return JSON.stringify({
      status: "ERRO",
      action,
      mensagem: erro instanceof Error ? erro.message : String(erro)
    });
  }
}

function obterVersao(workbook: ExcelScript.Workbook): string {
  const ws = workbook.getWorksheet("_sp_aguas_meta");
  if (!ws) return "";
  const valor = ws.getRange("A1").getValue();
  return valor === null || valor === undefined ? "" : String(valor);
}

function definirVersao(workbook: ExcelScript.Workbook, versao: string) {
  const ws = workbook.getWorksheet("_sp_aguas_meta") || workbook.addWorksheet("_sp_aguas_meta");
  ws.getRange("A1").setValue(versao);
  ws.setVisibility(ExcelScript.SheetVisibility.hidden);
}

function obterOuCriarAba(workbook: ExcelScript.Workbook, nome: string): ExcelScript.Worksheet {
  let ws = workbook.getWorksheet(nome);
  if (!ws) ws = workbook.addWorksheet(nome);
  return ws;
}

function lerSnapshot(workbook: ExcelScript.Workbook): Record<string, unknown[]> {
  const snapshot: Record<string, unknown[]> = {};

  for (const nome of Object.keys(COLUNAS)) {
    const ws = obterOuCriarAba(workbook, nome);
    const usado = ws.getUsedRange();
    if (!usado) {
      snapshot[nome] = [];
      continue;
    }

    const valores = usado.getValues();
    if (valores.length <= 1) {
      snapshot[nome] = [];
      continue;
    }

    const cabecalho = valores[0].map(v => String(v ?? "").trim());
    const indices = COLUNAS[nome].map(coluna => {
      const idx = cabecalho.indexOf(coluna);
      if (idx < 0) throw new Error(`A aba ${nome} não possui a coluna ${coluna}.`);
      return idx;
    });

    snapshot[nome] = valores.slice(1)
      .filter(linha => linha.some(v => v !== "" && v !== null && v !== undefined))
      .map(linha => {
        const registro: Record<string, unknown> = {};
        COLUNAS[nome].forEach((coluna, i) => {
          registro[coluna] = normalizarValor(linha[indices[i]]);
        });
        return registro;
      });
  }

  return snapshot;
}

function normalizarValor(valor: unknown): unknown {
  if (valor === null || valor === undefined) return "";
  return valor;
}

function validarSnapshot(snapshot: Record<string, unknown[]>) {
  for (const nome of Object.keys(COLUNAS)) {
    if (!Array.isArray(snapshot[nome])) {
      throw new Error(`Snapshot inválido: ${nome} deve ser uma lista.`);
    }
    for (const registro of snapshot[nome] as Record<string, unknown>[]) {
      for (const coluna of COLUNAS[nome]) {
        if (!(coluna in registro)) {
          throw new Error(`Snapshot inválido: ${nome}.${coluna} está ausente.`);
        }
      }
    }
  }
}

function gravarSnapshot(workbook: ExcelScript.Workbook, snapshot: Record<string, unknown[]>) {
  for (const nome of Object.keys(COLUNAS)) {
    const ws = obterOuCriarAba(workbook, nome);
    const usado = ws.getUsedRange();
    if (usado) {
      const linhas = usado.getRowCount();
      const colunas = usado.getColumnCount();
      if (linhas > 1) {
        ws.getRangeByIndexes(1, 0, linhas - 1, colunas).clear(ExcelScript.ClearApplyTo.contents);
      }
    }

    ws.getRangeByIndexes(0, 0, 1, COLUNAS[nome].length).setValues([COLUNAS[nome]]);

    const registros = snapshot[nome] as Record<string, unknown>[];
    if (registros.length === 0) continue;

    const matriz = registros.map(registro => COLUNAS[nome].map(coluna => valorParaExcel(registro[coluna])));
    ws.getRangeByIndexes(1, 0, matriz.length, COLUNAS[nome].length).setValues(matriz);
  }
}

function valorParaExcel(valor: unknown): string | number | boolean {
  if (valor === null || valor === undefined) return "";
  if (typeof valor === "number" || typeof valor === "boolean") return valor;
  return String(valor);
}
