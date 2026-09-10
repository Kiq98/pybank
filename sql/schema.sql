CREATE TABLE bancos(
    id SERIAL PRIMARY KEY,
    nome TEXT NOT NULL,
    UNIQUE(nome)
);

CREATE TABLE faturas(
    id SERIAL PRIMARY KEY,
    banco_id INT NOT NULL,
    FOREIGN KEY (banco_id) REFERENCES bancos(id),
    ano INT NOT NULL,
    mes INT NOT NULL,
    total NUMERIC(10,2),
    feedback TEXT,
    UNIQUE(banco_id, ano, mes)
);

CREATE TABLE transacoes(
    id SERIAL PRIMARY KEY,
    fatura_id INT NOT NULL,
    FOREIGN KEY (fatura_id) REFERENCES faturas(id) ON DELETE CASCADE,
    titular TEXT,
    data DATE,
    descricao TEXT NOT NULL,
    valor NUMERIC(10,2) NOT NULL,
    parcela TEXT
);