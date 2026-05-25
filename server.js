// ===================================================
// BACKEND - INTRANET CCO SLU
// ===================================================

const express = require("express");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");
const sqlite3 = require("sqlite3").verbose();
const path = require("path");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 3000;

// Permite comunicação com o frontend
app.use(cors());
app.use(express.json());

// Pasta temporária para uploads
const upload = multer({ dest: "Uploads/" });

// Banco SQLite
const dbPath = path.join(__dirname, "Database", "cco.db");
const db = new sqlite3.Database(dbPath);

// Cria tabela se não existir
db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS dados_operacionais (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      servico TEXT,
      data TEXT,
      quantidade REAL,
      unidade TEXT,
      valor_unitario REAL,
      valor_total REAL,
      origem TEXT,
      criado_em DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  `);
});

// ===================================================
// ROTA TESTE
// ===================================================
app.get("/", (req, res) => {
  res.json({ mensagem: "Backend CCO funcionando" });
});

// ===================================================
// LISTAR DADOS PARA TODOS OS USUÁRIOS
// ===================================================
app.get("/dados", (req, res) => {
  db.all(
    "SELECT * FROM dados_operacionais ORDER BY data DESC",
    [],
    (erro, linhas) => {
      if (erro) {
        return res.status(500).json({ erro: "Erro ao buscar dados" });
      }

      res.json(linhas);
    }
  );
});

// ===================================================
// IMPORTAR PLANILHA
// Admin e CCO importam.
// Os dados ficam salvos no banco e visíveis para todos.
// ===================================================
app.post("/importar", upload.single("arquivo"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ erro: "Nenhum arquivo enviado" });
    }

    const workbook = XLSX.readFile(req.file.path);
    const nomeAba = workbook.SheetNames[0];
    const planilha = workbook.Sheets[nomeAba];

    const dados = XLSX.utils.sheet_to_json(planilha, {
      defval: ""
    });

    const stmt = db.prepare(`
      INSERT INTO dados_operacionais
      (servico, data, quantidade, unidade, valor_unitario, valor_total, origem)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `);

    dados.forEach((linha) => {
      const servico =
        linha.Serviço ||
        linha.Servico ||
        linha.SERVIÇO ||
        linha.SERVICO ||
        linha.servico ||
        "Não informado";

      const data =
        linha.Data ||
        linha.DATA ||
        linha.data ||
        "";

      const quantidade =
        Number(
          linha.Quantidade ||
          linha.QTD ||
          linha.qtd ||
          linha.quantidade ||
          0
        );

      const unidade =
        linha.Unidade ||
        linha.UNIDADE ||
        linha.unidade ||
        "";

      const valorUnitario =
        Number(
          linha.ValorUnitario ||
          linha["Valor Unitário"] ||
          linha.valor_unitario ||
          0
        );

      const valorTotal =
        Number(
          linha.ValorTotal ||
          linha["Valor Total"] ||
          linha.valor_total ||
          quantidade * valorUnitario ||
          0
        );

      stmt.run(
        servico,
        data,
        quantidade,
        unidade,
        valorUnitario,
        valorTotal,
        req.file.originalname
      );
    });

    stmt.finalize();

    res.json({
      mensagem: "Planilha importada com sucesso",
      totalImportado: dados.length
    });

  } catch (erro) {
    res.status(500).json({
      erro: "Erro ao importar planilha",
      detalhe: erro.message
    });
  }
});

// ===================================================
// LIMPAR DADOS
// Apenas Admin deve usar no frontend
// ===================================================
app.delete("/dados", (req, res) => {
  db.run("DELETE FROM dados_operacionais", [], function (erro) {
    if (erro) {
      return res.status(500).json({ erro: "Erro ao limpar dados" });
    }

    res.json({ mensagem: "Dados apagados com sucesso" });
  });
});

// ===================================================
// INICIAR SERVIDOR
// ===================================================
app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});