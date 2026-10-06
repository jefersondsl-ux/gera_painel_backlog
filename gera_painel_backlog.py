#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gera_painel_backlog.py
----------------------
Gera o painel HTML de cobranca do backlog de instalacao de circuitos.

O que faz, na ordem:
  1. Le a base tratada (BD_Backlog_SGP.xlsx).
  2. Calcula as metricas do dia e grava um SNAPSHOT em ./historico/.
  3. Procura o snapshot anterior e calcula os DELTAS (dia a dia ou semana a semana).
  4. Le ./anotacoes.json (cobrancas registradas pelo gerente na rodada anterior).
  5. Gera um HTML autocontido, que abre offline e pode ser mandado por e-mail.

Uso:
    python gera_painel_backlog.py BD_Backlog_SGP.xlsx

O historico e o que torna a comparacao possivel. Rodar o script todo dia (ou
toda segunda) e o unico requisito: cada execucao deixa um snapshot no disco, e
o proximo painel compara contra ele automaticamente.
"""

import json
import re
import sys
from datetime import datetime
from pathlib import Path
import os
import shutil
import pandas as pd

# --------------------------------------------------------------------------
# Configuracao
# --------------------------------------------------------------------------

SUBCAMINHO_PADRAO = "BASES/Projeto_Cloude/Relatorios_Semanais_Backlog"


def caminho_base_claro(subcaminho):
    """Resolve a pasta de destino, com fallback em vez de travar o script.

    Ordem de prioridade:
      1. Variável de ambiente BACKLOG_BASE_DIR (caminho absoluto explícito).
      2. OneDrive corporativo (OneDriveCommercial) ou pessoal (OneDrive).
      3. Fallback local: pasta 'Relatorios_Semanais_Backlog' ao lado do script.
         (Isso evita que o script quebre em tarefas agendadas ou ambientes
         onde o OneDrive não está sincronizado — só avisa no console.)
    """
    override = os.environ.get("BACKLOG_BASE_DIR")
    if override:
        print(f"Usando BACKLOG_BASE_DIR definido no ambiente: {override}")
        return Path(override)

    onedrive_path = os.environ.get("OneDriveCommercial") or os.environ.get("OneDrive")
    if onedrive_path:
        return Path(onedrive_path, *subcaminho.split('/'))

    fallback = Path(__file__).resolve().parent / "Relatorios_Semanais_Backlog"
    print("AVISO: variável de ambiente OneDrive/OneDriveCommercial não encontrada "
          "(comum em tarefas agendadas que não herdam o ambiente do usuário).")
    print(f"Gravando localmente em: {fallback}")
    print("Para apontar para outra pasta, defina BACKLOG_BASE_DIR ou passe o "
          "caminho como 2º argumento na linha de comando.")
    return fallback


def resolver_base_dir():
    # 2º argumento de linha de comando tem prioridade máxima, se informado.
    if len(sys.argv) >= 3:
        return Path(sys.argv[2])
    return caminho_base_claro(SUBCAMINHO_PADRAO)


BASE_DIR = resolver_base_dir()
BASE_DIR.mkdir(parents=True, exist_ok=True)

DIR_HISTORICO = BASE_DIR / "historico"
ARQ_ANOTACOES = BASE_DIR / "anotacoes.json"
DIR_SAIDA = BASE_DIR

COL_AGING = "AGING_TAREFA_ATUAL"
COL_PROJ = "CARIMBO_PREFIXO"  # chave de projeto: usar prefixo do carimbo
COL_CARIMBO = "CARIMBO_PREFIXO"
COL_CIR = "COD_CIR"

print("BASE_DIR =", BASE_DIR)

# Quem e cobrado: o coordenador tecnico.
COL_COORD = "GER_TEC_AJUST"
# Por que esta parado: o tipo de pendencia (Cliente, Vendas, Projetos Especiais...).
# NAO e coordenacao — e a natureza do bloqueio.
COL_PENDENCIA = "TAREFA_RESPONSAVEL"


# --------------------------------------------------------------------------
# Leitura e metricas
# --------------------------------------------------------------------------
def carregar_base(caminho):
    try:
        df = pd.read_excel(caminho)
    except PermissionError:
        # Tenta ler por uma cópia temporária para contornar lock de aplicativos.
        origem = Path(caminho)
        copia_tmp = origem.with_name(f"~tmp_leitura_{origem.name}")
        try:
            shutil.copy2(origem, copia_tmp)
            df = pd.read_excel(copia_tmp)
        except PermissionError:
            raise SystemExit(
                "ERRO: sem permissão para ler a base. Feche o arquivo no Excel, "
                "aguarde o OneDrive sincronizar e tente novamente."
            )
        finally:
            if copia_tmp.exists():
                try:
                    copia_tmp.unlink()
                except OSError:
                    pass

    faltando = [c for c in (COL_AGING, COL_PROJ, COL_CIR) if c not in df.columns]
    if faltando:
        raise SystemExit(f"ERRO: colunas ausentes na base: {faltando}")
    return df


def faixa_aging(v):
    if pd.isna(v):
        return "sem_dado"
    if v <= 15:
        return "ate_15"
    if v <= 30:
        return "d16_30"
    if v <= 60:
        return "d31_60"
    return "d60_mais"


def sla_dias(txt):
    try:
        return int(str(txt).split()[0])
    except (ValueError, IndexError, AttributeError):
        return None


def localizar_coluna_receita(df):
    for col in df.columns:
        if str(col).strip().upper() == "DELTA_RECEITA":
            return col
    return None


def parse_valor_monetario(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if not texto:
        return None
    texto = texto.replace("R$", "").replace("r$", "").replace("$", "")
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def normalizar_projeto(v):
    if pd.isna(v):
        return None
    if isinstance(v, str):
        s = v.strip()
        if not s or s.upper() == "NAN":
            return None
        return s
    return v


def contar_texto_exato(serie, texto_alvo):
    if serie is None:
        return 0
    normalizada = serie.fillna("").astype(str).str.strip().str.casefold()
    return int((normalizada == str(texto_alvo).strip().casefold()).sum())


def somar_receita_por_texto(df, coluna, texto_alvo):
  if coluna not in df.columns or "_delta_receita" not in df.columns:
    return 0
  filtro = df[coluna].fillna("").astype(str).str.strip().str.casefold()
  receita = df.loc[filtro == str(texto_alvo).strip().casefold(), "_delta_receita"]
  return round(float(receita.sum()), 2)


def contar_sem_carimbo_mais_3_dias(df):
    col_status = "CARIMBO_STATUS"
    col_aging = "AGING_BACKLOG"
    if col_status not in df.columns or col_aging not in df.columns:
        return 0
    sem_carimbo = df[col_status].fillna("").astype(str).str.strip().str.casefold()
    aging = pd.to_numeric(df[col_aging], errors="coerce")
    return int(((sem_carimbo == "sem carimbo") & (aging > 3)).sum())


_ORDEM_MESES_FORECAST = {
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
    "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}


def _prioridade_faixa_backlog(label):
    # Ordem fixa: Backlog Atual, meses (cronológico), Meses Restantes, A Ajustar, A Definir, Outros.
    if label == "BACKLOG_ATUAL":
        return (0, 0)
    m = re.match(r"^([A-Z]{3})_(\d{4})$", label)
    if m:
        mes, ano = m.groups()
        return (1, int(ano) * 12 + _ORDEM_MESES_FORECAST.get(mes, 0))
    if label == "MESES_RESTANTES":
        return (2, 0)
    if label == "AJUSTAR":
        return (3, 0)
    if label == "DEFINIR":
        return (4, 0)
    if label == "OUTROS":
        return (5, 0)
    return (6, label)


def _exibir_faixa_backlog(label):
    mapa = {
        "DEFINIR": "A Definir",
        "AJUSTAR": "A Ajustar",
        "BACKLOG_ATUAL": "Backlog Atual",
        "MESES_RESTANTES": "Meses Restantes",
        "OUTROS": "Outros",
    }
    if label in mapa:
        return mapa[label]
    return label.replace("_", " ").title()


def calcular_metricas(df):
    """Devolve o dicionario que vira snapshot. Tudo aqui e comparavel no tempo."""
    com_aging = df[df[COL_AGING].notna()]

    df = df.copy()
    if COL_PROJ in df.columns:
        df[COL_PROJ] = df[COL_PROJ].apply(normalizar_projeto)
    receita_col = localizar_coluna_receita(df)
    if receita_col is not None:
        df["_delta_receita"] = df[receita_col].map(parse_valor_monetario)
    for col, alias in [
            ("GROSS", "_gross"),
            ("FLAG_SERVICO", "_servico"),
            ("OUTROS", "_outros"),
            ("PRODUTO_INTERNET", "_internet"),
            ("PRODUTO_DADOS", "_dados"),
            ("PRODUTO_VOZ", "_voz"),
            ("PRODUTO_WIFI", "_wifi"),
        ]:
        if col in df.columns:
            df[alias] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "FAIXA_BACKLOG" in df.columns:
        df["_faixa_backlog_norm"] = df["FAIXA_BACKLOG"].fillna("DEFINIR").astype(str).str.strip().str.upper()
    df["_sla_d"] = df.get("SLA_TAREFA_ATUAL", pd.Series(dtype=object)).map(sla_dias)
    aferivel = df[df[COL_AGING].notna() & df["_sla_d"].notna()]
    fora_sla = int((aferivel[COL_AGING] > aferivel["_sla_d"]).sum())
    qtd_data_prevista_a_definir = contar_texto_exato(
      df.get("DATA_PREVISAO_ATIVACAO_CLIENTE_AJUSTADA", pd.Series(dtype=object)),
      "A Definir",
    )
    qtd_data_compromisso_sem_data = contar_texto_exato(
      df.get("DATA_COMPROMISSO_CLIENTE_Py", pd.Series(dtype=object)),
      "Sem Data",
    )
    qtd_pend_cliente = contar_texto_exato(
      df.get("TAREFA_RESPONSAVEL", pd.Series(dtype=object)),
      "Cliente",
    )
    qtd_pend_vendas = contar_texto_exato(
      df.get("TAREFA_RESPONSAVEL", pd.Series(dtype=object)),
      "Vendas",
    )
    qtd_pend_pje = contar_texto_exato(
      df.get("TAREFA_RESPONSAVEL", pd.Series(dtype=object)),
      "Projetos Especiais",
    )
    qtd_sem_carimbo_mais_3_dias = contar_sem_carimbo_mais_3_dias(df)

    faixas = com_aging[COL_AGING].map(faixa_aging).value_counts().to_dict()

    forecast = []
    if "_faixa_backlog_norm" in df.columns:
        contagem_forecast = df["_faixa_backlog_norm"].value_counts().to_dict()
        receita_forecast = (
            df.groupby(df["_faixa_backlog_norm"])["_delta_receita"].sum().to_dict()
            if "_delta_receita" in df.columns else {}
        )
        for label in sorted(contagem_forecast, key=_prioridade_faixa_backlog):
            forecast.append({
                "id": label,
                "tipo": _exibir_faixa_backlog(label),
                "circuitos": int(contagem_forecast[label]),
                "receita": round(float(receita_forecast.get(label, 0)), 2),
            })

    kpis = {
        "circuitos": int(len(df)),
        "clientes": int(df["CLIENTE"].dropna().nunique()) if "CLIENTE" in df else 0,
        "projetos": int(df[COL_PROJ].dropna().nunique()),
        "carimbos": int(df[COL_CARIMBO].nunique()) if COL_CARIMBO in df else 0,
        "entregues": int((df.get("STATUS_ACESSOS", pd.Series(dtype=object)) == "Entregue").sum()),
        "pendentes": int((df.get("STATUS_ACESSOS", pd.Series(dtype=object)) == "Pendente").sum()),
        "regra_comercial": int(df.get("REGRA_COMERCIAL_PY", pd.Series(dtype=int)).sum()),
        "estrategia_redes": int((df.get("ESTRATEGIA_REDES_PY", pd.Series(dtype=object)) == "SIM").sum()),
        "com_gargalo": int(df.get("FLAG_GARGALO", pd.Series(dtype=int)).sum()),
        "gross": int(df.get("_gross", pd.Series(dtype=int)).sum()) if "_gross" in df.columns else 0,
        "servicos": int(df.get("_servico", pd.Series(dtype=int)).sum()) if "_servico" in df.columns else 0,
        "outros": int(df.get("_outros", pd.Series(dtype=int)).sum()) if "_outros" in df.columns else 0,
        "internet": int(df.get("_internet", pd.Series(dtype=int)).sum()) if "_internet" in df.columns else 0,
        "dados": int(df.get("_dados", pd.Series(dtype=int)).sum()) if "_dados" in df.columns else 0,
        "voz": int(df.get("_voz", pd.Series(dtype=int)).sum()) if "_voz" in df.columns else 0,
        "wifi": int(df.get("_wifi", pd.Series(dtype=int)).sum()) if "_wifi" in df.columns else 0,
        "receita_total": round(float(df.get("_delta_receita", pd.Series(dtype=float)).sum()), 2) if "_delta_receita" in df.columns else 0,
        "receita_gross": round(float(df.loc[df.get("_gross", 0) > 0, "_delta_receita"].sum()), 2) if "_delta_receita" in df.columns and "_gross" in df.columns else 0,
        "receita_servicos": round(float(df.loc[df.get("_servico", 0) > 0, "_delta_receita"].sum()), 2) if "_delta_receita" in df.columns and "_servico" in df.columns else 0,
        "receita_outros": round(float(df.loc[df.get("_outros", 0) > 0, "_delta_receita"].sum()), 2) if "_delta_receita" in df.columns and "_outros" in df.columns else 0,
        "receita_internet": round(float(df.loc[df.get("_internet", 0) > 0, "_delta_receita"].sum()), 2) if "_delta_receita" in df.columns and "_internet" in df.columns else 0,
        "receita_dados": round(float(df.loc[df.get("_dados", 0) > 0, "_delta_receita"].sum()), 2) if "_delta_receita" in df.columns and "_dados" in df.columns else 0,
        "receita_voz": round(float(df.loc[df.get("_voz", 0) > 0, "_delta_receita"].sum()), 2) if "_delta_receita" in df.columns and "_voz" in df.columns else 0,
        "receita_wifi": round(float(df.loc[df.get("_wifi", 0) > 0, "_delta_receita"].sum()), 2) if "_delta_receita" in df.columns and "_wifi" in df.columns else 0,
        "receita_pendentes": round(float(df.loc[df.get("STATUS_ACESSOS", pd.Series(dtype=object)) == "Pendente", "_delta_receita"].sum()), 2) if "_delta_receita" in df.columns else 0,
        "receita_entregues": round(float(df.loc[df.get("STATUS_ACESSOS", pd.Series(dtype=object)) == "Entregue", "_delta_receita"].sum()), 2) if "_delta_receita" in df.columns else 0,
        "receita_pend_cliente": somar_receita_por_texto(df, COL_PENDENCIA, "Cliente"),
        "receita_pend_vendas": somar_receita_por_texto(df, COL_PENDENCIA, "Vendas"),
        "receita_pend_pje": somar_receita_por_texto(df, COL_PENDENCIA, "Projetos Especiais"),
        "sem_carimbo": int(df[COL_CARIMBO].isna().sum()) if COL_CARIMBO in df else 0,
        "sem_carimbo_mais_3_dias": qtd_sem_carimbo_mais_3_dias,
        "sem_aging": int(df[COL_AGING].isna().sum()),
        "data_prevista_ativacao_a_definir": qtd_data_prevista_a_definir,
        "data_compromisso_sem_data": qtd_data_compromisso_sem_data,
        "datas_nao_preenchidas": qtd_data_prevista_a_definir + qtd_data_compromisso_sem_data,
        "pend_cliente": qtd_pend_cliente,
        "pend_vendas": qtd_pend_vendas,
        "pend_pje": qtd_pend_pje,
        "fora_sla": fora_sla,
        "aferivel_sla": int(len(aferivel)),
        "aging_medio": round(float(com_aging[COL_AGING].mean()), 1) if len(com_aging) else 0,
        "aging_max": int(com_aging[COL_AGING].max()) if len(com_aging) else 0,
        "ate_15": int(faixas.get("ate_15", 0)),
        "d16_30": int(faixas.get("d16_30", 0)),
        "d31_60": int(faixas.get("d31_60", 0)),
        "d60_mais": int(faixas.get("d60_mais", 0)),
    }

    # Frentes de cobranca — uma por COORDENADOR (GER_TEC_AJUST).
    # Dentro de cada uma, a quebra por tipo de pendencia, que e o argumento da cobranca.
    frentes = []
    if COL_COORD in df.columns:
        agg = {
            "circuitos": (COL_CIR, "count"),
            "projetos": (COL_PROJ, lambda s: int(s.dropna().nunique())),
            "gargalo": ("FLAG_GARGALO", "sum"),
            "aging_medio": (COL_AGING, "mean"),
            "aging_max": (COL_AGING, "max"),
            "criticos": (COL_AGING, lambda s: int((s > 60).sum())),
            "sem_aging": (COL_AGING, lambda s: int(s.isna().sum())),
        }
        if "_delta_receita" in df.columns:
            agg["receita"] = ("_delta_receita", "sum")
        if "_gross" in df.columns:
            agg["gross"] = ("_gross", "sum")
        if "_servico" in df.columns:
            agg["servicos"] = ("_servico", "sum")
        if "_outros" in df.columns:
            agg["outros"] = ("_outros", "sum")
        if "_internet" in df.columns:
            agg["internet"] = ("_internet", "sum")
        if "_dados" in df.columns:
            agg["dados"] = ("_dados", "sum")
        if "_voz" in df.columns:
            agg["voz"] = ("_voz", "sum")
        if "_wifi" in df.columns:
            agg["wifi"] = ("_wifi", "sum")
        if "STATUS_ACESSOS" in df.columns:
            agg["pendentes"] = ("STATUS_ACESSOS", lambda s: int((s == "Pendente").sum()))
            agg["entregues"] = ("STATUS_ACESSOS", lambda s: int((s == "Entregue").sum()))
        if COL_PENDENCIA in df.columns:
          agg["pend_cliente"] = (COL_PENDENCIA, lambda s: int((s == "Cliente").sum()))
          agg["pend_vendas"] = (COL_PENDENCIA, lambda s: int((s == "Vendas").sum()))
          agg["pend_pje"] = (COL_PENDENCIA, lambda s: int((s == "Projetos Especiais").sum()))
        if COL_CARIMBO in df.columns:
            agg["sem_carimbo"] = (COL_CARIMBO, lambda s: int(s.isna().sum()))
        if "DATA_PREVISAO_ATIVACAO_CLIENTE_AJUSTADA" in df.columns:
            agg["data_prevista_ativacao"] = (
                "DATA_PREVISAO_ATIVACAO_CLIENTE_AJUSTADA",
                lambda s: int((s.fillna("").astype(str).str.strip().str.casefold() == "a definir").sum()),
            )
        if "DATA_COMPROMISSO_CLIENTE_Py" in df.columns:
            agg["data_compromisso_cliente"] = (
                "DATA_COMPROMISSO_CLIENTE_Py",
                lambda s: int((s.fillna("").astype(str).str.strip().str.casefold() == "sem data").sum()),
            )
        g = df.groupby(COL_COORD).agg(**agg)
        for nome, r in g.sort_values("circuitos", ascending=False).iterrows():
            sub = df[df[COL_COORD] == nome]

            forecast_por_frente = {}
            if "_faixa_backlog_norm" in df.columns:
                contagem_sub_forecast = sub["_faixa_backlog_norm"].value_counts().to_dict()
                for item in forecast:
                    forecast_por_frente[item["id"]] = int(contagem_sub_forecast.get(item["id"], 0))

            pendencias = []
            if COL_PENDENCIA in df.columns:
                p = sub.groupby(COL_PENDENCIA).agg(
                    circuitos=(COL_CIR, "count"),
                    aging_medio=(COL_AGING, "mean"),
                )
                for pn, pr in p.sort_values("circuitos", ascending=False).iterrows():
                    pendencias.append({
                        "tipo": str(pn),
                        "circuitos": int(pr.circuitos),
                        "aging_medio": (round(float(pr.aging_medio), 1)
                                        if pd.notna(pr.aging_medio) else None),
                    })

            # Os 3 circuitos mais antigos sob este coordenador — a cobranca nominal
            piores = []
            sc = sub[sub[COL_AGING].notna()].nlargest(3, COL_AGING)
            for _, cr in sc.iterrows():
                piores.append({
                    "designacao": str(cr.get("DESIGNACAO", "")),
                    "cliente": str(cr.get("CLIENTE_CANON", "")),
                    "carimbo": ("" if pd.isna(cr.get(COL_CARIMBO))
                                else str(cr.get(COL_CARIMBO))),
                    "aging": int(cr[COL_AGING]),
                    "tarefa": str(cr.get("TAREFA_ATUAL", "")),
                })

            frente = {
                "id": str(nome),
                "coordenador": str(nome),
                "circuitos": int(r.circuitos),
              "circuitos_ids": sorted(int(x) for x in sub[COL_CIR].dropna().unique()),
                "forecast": forecast_por_frente,
                "projetos": int(r.projetos),
                "pct_gargalo": round(float(r.gargalo) / float(r.circuitos) * 100, 1),
                "aging_medio": round(float(r.aging_medio), 1) if pd.notna(r.aging_medio) else 0,
                "aging_max": int(r.aging_max) if pd.notna(r.aging_max) else 0,
                "criticos": int(r.criticos),
                "sem_aging": int(r.sem_aging),
                "pendencias": pendencias,
                "piores": piores,
            }
            if "receita" in r.index and pd.notna(r.receita):
                frente["receita"] = round(float(r.receita), 2)
            if "gross" in r.index:
                frente["gross"] = int(r.gross)
            if "servicos" in r.index:
                frente["servicos"] = int(r.servicos)
            if "outros" in r.index:
                frente["outros"] = int(r.outros)
            if "internet" in r.index:
                frente["internet"] = int(r.internet)
            if "dados" in r.index:
                frente["dados"] = int(r.dados)
            if "voz" in r.index:
                frente["voz"] = int(r.voz)
            if "wifi" in r.index:
                frente["wifi"] = int(r.wifi)
            if "pendentes" in r.index:
                frente["pendentes"] = int(r.pendentes)
            if "entregues" in r.index:
                frente["entregues"] = int(r.entregues)
            if "pend_cliente" in r.index:
              frente["pend_cliente"] = int(r.pend_cliente)
            if "pend_vendas" in r.index:
              frente["pend_vendas"] = int(r.pend_vendas)
            if "pend_pje" in r.index:
              frente["pend_pje"] = int(r.pend_pje)
            if "sem_carimbo" in r.index:
                frente["sem_carimbo"] = int(r.sem_carimbo)
            if "data_prevista_ativacao" in r.index:
                frente["data_prevista_ativacao"] = int(r.data_prevista_ativacao)
            if "data_compromisso_cliente" in r.index:
                frente["data_compromisso_cliente"] = int(r.data_compromisso_cliente)
            frentes.append(frente)

    # Pendencias no total geral — visao transversal, nao e cobranca de ninguem
    pendencias_geral = []
    if COL_PENDENCIA in df.columns:
        g = df.groupby(COL_PENDENCIA).agg(
            circuitos=(COL_CIR, "count"),
            projetos=(COL_PROJ, lambda s: int(s.dropna().nunique())),
            aging_medio=(COL_AGING, "mean"),
            criticos=(COL_AGING, lambda s: int((s > 60).sum())),
        )
        for nome, r in g.sort_values("circuitos", ascending=False).iterrows():
            pendencias_geral.append({
                "id": str(nome),
                "tipo": str(nome),
                "circuitos": int(r.circuitos),
                "projetos": int(r.projetos),
                "aging_medio": round(float(r.aging_medio), 1) if pd.notna(r.aging_medio) else None,
                "criticos": int(r.criticos),
            })

    # Gargalos por tarefa, ordenados por impacto (circuitos x aging medio)
    gargalos = []
    if "TAREFA_ATUAL" in df.columns:
        g = com_aging.groupby("TAREFA_ATUAL").agg(
            circuitos=(COL_CIR, "count"),
            projetos=(COL_PROJ, lambda s: int(s.dropna().nunique())),
            aging_medio=(COL_AGING, "mean"),
            aging_max=(COL_AGING, "max"),
        )
        g["impacto"] = g.circuitos * g.aging_medio
        for nome, r in g.sort_values("impacto", ascending=False).head(10).iterrows():
            gargalos.append({
                "id": str(nome).split(" - ")[0].strip(),
                "tarefa": str(nome),
                "circuitos": int(r.circuitos),
                "projetos": int(r.projetos),
                "aging_medio": round(float(r.aging_medio), 1),
                "aging_max": int(r.aging_max),
            })

    # Clientes
    clientes = []
    if "CLIENTE_CANON" in df.columns:
        g = df.groupby("CLIENTE_CANON").agg(
            circuitos=(COL_CIR, "count"),
            projetos=(COL_PROJ, lambda s: int(s.dropna().nunique())),
            gargalo=("FLAG_GARGALO", "sum"),
            aging_medio=(COL_AGING, "mean"),
        )
        for nome, r in g.sort_values("circuitos", ascending=False).head(12).iterrows():
            clientes.append({
                "id": str(nome),
                "cliente": str(nome),
                "circuitos": int(r.circuitos),
                "projetos": int(r.projetos),
                "pct_gargalo": round(float(r.gargalo) / float(r.circuitos) * 100, 1),
                "aging_medio": round(float(r.aging_medio), 1) if pd.notna(r.aging_medio) else None,
            })

    data_report = None
    if "DATA_REPORT" in df.columns and df["DATA_REPORT"].notna().any():
        data_report = pd.to_datetime(df["DATA_REPORT"]).max().strftime("%Y-%m-%d")

    return {
        "data_report": data_report or datetime.now().strftime("%Y-%m-%d"),
        "gerado_em": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "kpis": kpis,
        "frentes": frentes,
        "pendencias_geral": pendencias_geral,
        "gargalos": gargalos,
        "clientes": clientes,
        "forecast": forecast,
        # lista de circuitos: permite calcular entradas e saidas entre snapshots
        "circuitos_ids": sorted(int(x) for x in df[COL_CIR].dropna().unique()),
        "circuitos_ids_gross": (sorted(int(x) for x in df.loc[df["_gross"] > 0, COL_CIR].dropna().unique())
                    if "_gross" in df.columns else []),
        "circuitos_ids_servicos": (sorted(int(x) for x in df.loc[df["_servico"] > 0, COL_CIR].dropna().unique())
                       if "_servico" in df.columns else []),
        "circuitos_ids_outros": (sorted(int(x) for x in df.loc[df["_outros"] > 0, COL_CIR].dropna().unique())
                     if "_outros" in df.columns else []),
        "circuitos_ids_internet": (sorted(int(x) for x in df.loc[df["_internet"] > 0, COL_CIR].dropna().unique())
                 if "_internet" in df.columns else []),
        "circuitos_ids_dados": (sorted(int(x) for x in df.loc[df["_dados"] > 0, COL_CIR].dropna().unique())
              if "_dados" in df.columns else []),
        "circuitos_ids_voz": (sorted(int(x) for x in df.loc[df["_voz"] > 0, COL_CIR].dropna().unique())
            if "_voz" in df.columns else []),
        "circuitos_ids_wifi": (sorted(int(x) for x in df.loc[df["_wifi"] > 0, COL_CIR].dropna().unique())
             if "_wifi" in df.columns else []),
    }


# --------------------------------------------------------------------------
# Historico e deltas
# --------------------------------------------------------------------------
def salvar_snapshot(snap):
    DIR_HISTORICO.mkdir(exist_ok=True)
    caminho = DIR_HISTORICO / f"snapshot_{snap['data_report']}.json"
    caminho.write_text(json.dumps(snap, ensure_ascii=False), encoding="utf-8")
    return caminho


def snapshot_anterior(data_atual):
    if not DIR_HISTORICO.exists():
        return None
    anteriores = sorted(
        p for p in DIR_HISTORICO.glob("snapshot_*.json")
        if p.stem.replace("snapshot_", "") < data_atual
    )
    if not anteriores:
        return None
    return json.loads(anteriores[-1].read_text(encoding="utf-8"))


def calcular_deltas(atual, anterior):
    """Compara dois snapshots. Devolve None se nao houver base de comparacao."""
    if not anterior:
        return None

    d = {"data_base": anterior["data_report"], "kpis": {}, "frentes": {},
         "pendencias_geral": {}, "gargalos": {}, "clientes": {}, "forecast": {}}

    for k, v in atual["kpis"].items():
        antes = anterior["kpis"].get(k)
        d["kpis"][k] = None if antes is None else round(v - antes, 1)

    def diff_lista(chave, campos):
        antes = {i["id"]: i for i in anterior.get(chave, [])}
        for item in atual.get(chave, []):
            a = antes.get(item["id"])
            if a:
                d[chave][item["id"]] = {c: round(item[c] - a[c], 1) for c in campos
                                        if item.get(c) is not None and a.get(c) is not None}
            else:
                d[chave][item["id"]] = {"novo": True}

    diff_lista("frentes", ["circuitos", "aging_medio", "criticos", "pct_gargalo", "receita"])

    # Entradas por gerente: compara IDs de circuito entre snapshots.
    antes_frentes = {i["id"]: i for i in anterior.get("frentes", [])}
    for item in atual.get("frentes", []):
      a = antes_frentes.get(item["id"])
      hoje_ids = set(item.get("circuitos_ids", []))

      # Só calcula entradas por gerente se houver base comparável por gerente.
      if a is None:
        entradas_gerente = len(hoje_ids)
      elif "circuitos_ids" in a:
        ontem_ids = set(a.get("circuitos_ids", []))
        entradas_gerente = len(hoje_ids - ontem_ids)
      else:
        entradas_gerente = None

      d["frentes"].setdefault(item["id"], {})["entradas"] = entradas_gerente

    diff_lista("pendencias_geral", ["circuitos", "aging_medio", "criticos"])
    diff_lista("gargalos", ["circuitos", "aging_medio"])
    diff_lista("clientes", ["circuitos", "pct_gargalo"])
    diff_lista("forecast", ["circuitos", "receita"])

    hoje = set(atual["circuitos_ids"])
    ontem = set(anterior["circuitos_ids"])
    entradas_ids = hoje - ontem
    d["entradas"] = len(entradas_ids)
    d["saidas"] = len(ontem - hoje)
    d["kpis_entradas"] = {
      "circuitos": len(entradas_ids),
      "gross": len(entradas_ids & set(atual.get("circuitos_ids_gross", []))),
      "servicos": len(entradas_ids & set(atual.get("circuitos_ids_servicos", []))),
      "outros": len(entradas_ids & set(atual.get("circuitos_ids_outros", []))),
      "internet": len(entradas_ids & set(atual.get("circuitos_ids_internet", []))),
      "dados": len(entradas_ids & set(atual.get("circuitos_ids_dados", []))),
      "voz": len(entradas_ids & set(atual.get("circuitos_ids_voz", []))),
      "wifi": len(entradas_ids & set(atual.get("circuitos_ids_wifi", []))),
    }
    return d


def carregar_anotacoes():
    if ARQ_ANOTACOES.exists():
        try:
            return json.loads(ARQ_ANOTACOES.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            print("AVISO: anotacoes.json invalido, ignorando.")
    return {}


# --------------------------------------------------------------------------
# HTML
# --------------------------------------------------------------------------
def gerar_html(snap, deltas, anotacoes):
    dados = json.dumps({"snap": snap, "deltas": deltas, "anotacoes": anotacoes},
                       ensure_ascii=False)
    template = HTML_TEMPLATE.replace("/*__DADOS__*/", dados)
    saida = DIR_SAIDA / f"painel_backlog_{snap['data_report']}.html"
    saida.write_text(template, encoding="utf-8")
    return saida


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Painel de Cobrança — Backlog de Circuitos</title>
<style>
  :root{
    --papel:#EBEEED; --papel-2:#F6F8F7; --tinta:#122029; --tinta-3:#5C6E75;
    --linha:#C3CCC9; --aqua:#0E8F8B; --ambar:#C4890A; --oxido:#B03F26;
    --realce:#FFF3C9;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--papel);color:var(--tinta);
    font:15px/1.55 "Segoe UI",-apple-system,Roboto,Helvetica,Arial,sans-serif}
  .mono{font-family:ui-monospace,"Cascadia Mono",Consolas,"DejaVu Sans Mono",monospace}
  .wrap{width:min(100%,1600px);max-width:1600px;margin:0 auto;padding:0 8px 16px}

  header{border-bottom:2px solid var(--tinta);padding:34px 0 18px;margin-bottom:6px}
  .eyebrow{font-size:11px;letter-spacing:.22em;text-transform:uppercase;
    color:var(--tinta-3);font-weight:600}
  h1{font-size:33px;line-height:1.1;margin:8px 0 14px;font-weight:700;letter-spacing:-.02em}
  .meta{display:flex;flex-wrap:wrap;gap:10px 26px;font-size:12.5px;color:var(--tinta-3)}
  .meta b{color:var(--tinta);font-weight:600}

  h2{font-size:12px;letter-spacing:.18em;text-transform:uppercase;font-weight:700;
    margin:46px 0 4px;padding-bottom:8px;border-bottom:1px solid var(--linha)}
  .sub{font-size:13.5px;color:var(--tinta-3);margin:10px 0 20px;max-width:62ch}

  /* KPIs */
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(158px,1fr));
    gap:6px;background:transparent;border:0;margin-top:20px}
  .kpi{position:relative;background:var(--papel-2);padding:15px 16px 15px 12px;border-left:6px solid transparent}
  .kpi.grp-base{border-left-color:#0E8F8B;background:#F2FAF9}
  .kpi.grp-classificacao{border-left-color:#B87410;background:#FDF5E8}
  .kpi.grp-produto{border-left-color:#2D6FA6;background:#EFF5FB}
  .kpi.grp-status{border-left-color:#6E9E5C;background:#F2F8EE}
  .kpi.grp-risco{border-left-color:#B03F26;background:#FCEFEB}
  .kpi.grp-pendencias{border-left-color:#B03F26;background:#FCEFEB}
  .kpi.grp-dados-ausentes{border-left-color:#7A8791;background:#F1F4F5}
  .kpi.grp-forecast{border-left-color:#D9720A;background:#FDF0DF}
  .kpi-pct{position:absolute;top:8px;right:10px;padding:2px 8px;border-radius:999px;
    font-size:14px;font-weight:700;letter-spacing:.02em;line-height:1.2}
  .kpi-pct.grp-base{background:#D3EAE8;color:#0A6E6B}
  .kpi-pct.grp-classificacao{background:#F3DDB9;color:#8B5908}
  .kpi-pct.grp-produto{background:#DCE8F7;color:#245B89}
  .kpi-pct.grp-status{background:#DDEDD2;color:#4E7A3F}
  .kpi-pct.grp-risco{background:#F3D6CE;color:#9A351F}
  .kpi-pct.grp-pendencias{background:#F3D6CE;color:#9A351F}
  .kpi-pct.grp-dados-ausentes{background:#DDE3E6;color:#5F6B73}
  .kpi-pct.grp-forecast{background:#F5D9B8;color:#8B4E08}
  .kpi-legenda{display:flex;flex-wrap:wrap;gap:14px;margin-top:10px;font-size:11.5px;color:var(--tinta-3)}
  .kpi-legenda span{display:flex;align-items:center;gap:6px}
  .kpi-legenda i{width:11px;height:11px;display:inline-block}
  .matriz{margin-top:24px;padding:16px;background:#fff;border:1px solid var(--linha);overflow-x:auto;}
  .matriz h3{margin:0 0 12px;font-size:14px;letter-spacing:.12em;text-transform:uppercase;color:var(--tinta-3)}
  .matriz-table{table-layout:fixed;min-width:1120px;border-collapse:collapse;font-size:13.2px}
  .matriz-table th,.matriz-table td{padding:9px 8px;border-bottom:1px solid var(--linha);overflow:hidden;}
  .matriz-table th{text-align:left;font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--tinta-3);font-weight:700}
  .matriz-table td.num{text-align:right}
  .matriz-table td.money{text-align:right;font-variant-numeric:tabular-nums}
  .matriz-table td.money,.matriz-table th.money{padding-left:10px;padding-right:10px;}
  .matriz-table thead tr.group-label th.group{text-align:center}
  .matriz-table thead tr.group-header th{text-align:center}
  .matriz-table th:not(:last-child),.matriz-table td:not(:last-child){border-right:1px solid rgba(195,204,201,.7)}
  .matriz-table th:first-child,.matriz-table td:first-child{position:sticky;left:0;
    background:#fff;box-shadow:3px 0 0 rgba(195,204,201,.8);z-index:2}
  .matriz-table thead th:first-child{z-index:4}
  .matriz-table thead tr.group-label th:first-child{background:#fff}
  .matriz-table tbody td:nth-child(5),.matriz-table tbody td:nth-child(8),
  .matriz-table tbody td:nth-child(12),.matriz-table tbody td:nth-child(14),
  .matriz-table tbody td:nth-child(17),
  .matriz-table thead tr.group-header th:nth-child(4),
  .matriz-table thead tr.group-header th:nth-child(7),
  .matriz-table thead tr.group-header th:nth-child(11),
  .matriz-table thead tr.group-header th:nth-child(13),
  .matriz-table thead tr.group-header th:nth-child(16),
  .matriz-table td.forecast-start,.matriz-table th.forecast-start{border-left:2px solid var(--tinta-3)}
  .matriz-bar{position:relative;display:block;overflow:hidden;text-align:right;font-variant-numeric:tabular-nums}
  .matriz-bar::before{content:"";position:absolute;inset:2px 0 2px auto;width:var(--bar-width);
    background:var(--bar-color);opacity:.16}
  .matriz-bar span{position:relative;z-index:1}
  .matriz-bar.backlog{--bar-color:#0E8F8B}
  .matriz-bar.classificacao{--bar-color:#B87410}
  .matriz-bar.produto{--bar-color:#2D6FA6}
  .matriz-bar.status{--bar-color:#6E9E5C}
  .matriz-bar.pendencias{--bar-color:#B03F26}
  .matriz-bar.dados-ausentes{--bar-color:#7A8791}
  .matriz-bar.forecast{--bar-color:#D9720A}
  .matriz-table tr.total td{font-weight:700;border-top:2px solid var(--linha);}
  .kpi .lbl{display:block;padding-right:78px;font-size:10.5px;letter-spacing:.11em;
    text-transform:uppercase;color:var(--tinta-3);font-weight:600;line-height:1.25}
  .kpi .val.val-composto + .none,.kpi .val.val-composto ~ .kpi-footer{display:none}
  .kpi.card-datas_nao_preenchidas .lbl{padding-right:0;white-space:nowrap;font-size:10px;letter-spacing:.08em}
  .kpi .val{font-size:29px;font-weight:700;letter-spacing:-.03em;margin-top:5px}
  .kpi .val small{font-size:14px;font-weight:600;color:var(--tinta-3)}
  .kpi .val.val-composto{font-size:15px;letter-spacing:0;margin-top:10px}
  .kpi-split{display:flex;flex-direction:column;gap:10px}
  .kpi.card-datas_nao_preenchidas .kpi-split{gap:24px}
  .kpi-split-item{display:flex;align-items:baseline;gap:7px;padding-top:6px;border-top:1px solid rgba(18,32,41,.18)}
  .kpi-split-item .qtd{display:inline-block;flex:0 0 38px;font-size:20px;font-weight:700;line-height:1.05;color:var(--tinta)}
  .kpi-split-item .txt{display:inline-block;margin-top:0;font-size:10px;line-height:1.2;
    letter-spacing:.04em;text-transform:uppercase;color:var(--tinta-3);white-space:nowrap}
  .kpi-split-item .txt{font-size:9px;letter-spacing:0;text-transform:none}
  .kpi-footer{display:block;margin-top:8px;font-size:12px;font-weight:600;color:#5C3A3A;letter-spacing:.02em;line-height:1.2}
  .kpi-footer-receita{margin-top:14px}
  .kpi-footer-right{text-align:right;color:var(--tinta-3)}
  .kpi-alerta{display:block;margin-top:9px;padding-top:7px;border-top:1px solid rgba(176,63,38,.28);
    color:var(--oxido);font-size:11px;font-weight:700;line-height:1.25}

  .delta{display:inline-block;font-size:11.5px;font-weight:700;margin-top:5px;
    padding:2px 7px;border-radius:2px;letter-spacing:.02em}
  .up{background:#F6DDD6;color:var(--oxido)}
  .down{background:#D3EAE8;color:var(--aqua)}
  .flat{background:#DDE3E1;color:var(--tinta-3)}
  .none{color:var(--tinta-3);font-size:11.5px;margin-top:5px;display:inline-block}

  /* Régua de aging */
  .regua{display:flex;height:34px;border:1px solid var(--linha);margin-top:18px}
  .regua div{display:flex;align-items:center;justify-content:center;
    font-size:11px;font-weight:700;color:#fff;overflow:hidden;white-space:nowrap}
  .legenda{display:flex;flex-wrap:wrap;gap:16px;margin-top:9px;font-size:11.5px;
    color:var(--tinta-3)}
  .legenda span{display:flex;align-items:center;gap:6px}
  .sw{width:11px;height:11px;display:inline-block}

  /* Frentes de cobrança */
  .frente{border:1px solid var(--linha);border-left:4px solid var(--tinta);
    background:var(--papel-2);margin-bottom:14px}
  .frente.crit{border-left-color:var(--oxido)}
  .frente.aten{border-left-color:var(--ambar)}
  .fhead{display:flex;flex-wrap:wrap;gap:12px;align-items:baseline;
    justify-content:space-between;padding:14px 18px;border-bottom:1px solid var(--linha)}
  .fnome{font-size:17px;font-weight:700;letter-spacing:-.01em}
  .fcoord{font-size:11px;letter-spacing:.13em;text-transform:uppercase;
    color:var(--tinta-3);font-weight:600;margin-top:2px}
  .fnums{display:flex;gap:22px;flex-wrap:wrap}
  .fnum{text-align:right}
  .fnum .n{font-size:20px;font-weight:700;letter-spacing:-.02em}
  .fnum .l{font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--tinta-3);
    font-weight:600}

  .fid{display:flex;gap:12px;align-items:center}
  .avatar{width:38px;height:38px;flex:0 0 38px;background:var(--tinta);color:var(--papel);
    display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;
    letter-spacing:.03em}
  .frente.crit .avatar{background:var(--oxido)}
  .frente.aten .avatar{background:var(--ambar)}

  .secao{padding:13px 18px;border-bottom:1px solid var(--linha)}
  .secao-t{font-size:10px;letter-spacing:.14em;text-transform:uppercase;
    color:var(--tinta-3);font-weight:700;margin-bottom:9px}
  .pends{display:grid;grid-template-columns:repeat(auto-fit,minmax(118px,1fr));gap:9px}
  .pend{border-left:2px solid var(--linha);padding-left:9px}
  .pend-t{font-size:11.5px;color:var(--tinta-3);line-height:1.25}
  .pend-n{font-size:18px;font-weight:700;letter-spacing:-.02em}
  .pend-a{font-size:11px;color:var(--tinta-3)}

  .piores{margin:0;padding:0;list-style:none;font-size:13px}
  .piores li{padding:5px 0;border-bottom:1px dotted var(--linha)}
  .piores li:last-child{border-bottom:0}
  .piores .mono{font-size:12.5px;background:#DFE5E3;padding:1px 5px}
  .piores .tar{display:block;color:var(--tinta-3);font-size:11.5px;margin-top:2px}

  /* Campo de cobrança */
  .cobranca{padding:14px 18px}
  .cobranca label{font-size:10.5px;letter-spacing:.13em;text-transform:uppercase;
    color:var(--tinta-3);font-weight:700;display:block;margin-bottom:7px}
  textarea{width:100%;min-height:62px;resize:vertical;padding:10px 12px;
    border:1px solid var(--linha);background:var(--realce);color:var(--tinta);
    font:14px/1.5 inherit;border-radius:0}
  textarea:focus{outline:2px solid var(--tinta);outline-offset:-2px;background:#fff}
  textarea::placeholder{color:#A2AEB2}
  .hist{font-size:12.5px;color:var(--tinta-3);margin-top:9px;padding-left:11px;
    border-left:2px solid var(--linha)}
  .hist b{color:var(--tinta)}

  table{width:100%;border-collapse:collapse;font-size:13.5px;margin-top:18px}
  th{text-align:left;font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;
    color:var(--tinta-3);font-weight:700;padding:9px 10px;border-bottom:2px solid var(--tinta)}
  td{padding:9px 10px;border-bottom:1px solid var(--linha)}
  th.num,td.num{text-align:right}
  tbody tr:hover{background:var(--papel-2)}
  .matriz-table thead tr.group-label th{font-size:12px;letter-spacing:.1em;text-transform:uppercase;font-weight:700;padding:10px;border-bottom:1px solid var(--linha)}
  .matriz-table thead tr.group-label th.group{background:#E8F6F5;color:var(--tinta);}
  .matriz-table thead tr.group-label th.classificacao{background:#F7ECE1;color:var(--tinta);}
  .matriz-table thead tr.group-label th.produto{background:#E9F0F7;color:var(--tinta);}
  .matriz-table thead tr.group-label th.status{background:#F2F8EE;color:var(--tinta);}
  .matriz-table thead tr.group-label th.pendencias{background:#FCEFEB;color:#9A351F;}
  .matriz-table thead tr.group-label th.dados-ausentes{background:#F1F4F5;color:var(--tinta);}
  .matriz-table thead tr.group-label th.forecast{background:#FDF0DF;color:var(--tinta);}
  .matriz-table thead tr.group-header th.backlog{background:#E8F6F5;}
  .matriz-table thead tr.group-header th.classificacao{background:#F7ECE1;}
  .matriz-table thead tr.group-header th.produto{background:#E9F0F7;}
  .matriz-table thead tr.group-header th.status{background:#F2F8EE;}
  .matriz-table thead tr.group-header th.pendencias{background:#FCEFEB;color:#9A351F;}
  .matriz-table thead tr.group-header th.dados-ausentes{background:#F1F4F5;}
  .matriz-table thead tr.group-header th.forecast{background:#FDF0DF;}
  .matriz-table thead tr.group-label th:first-child{background:transparent;border-bottom:none}
  .matriz-table thead tr.group-header th{border-bottom:2px solid var(--tinta)}

  .barra-bg{display:inline-block;width:62px;height:7px;background:var(--linha);
    vertical-align:middle;margin-right:7px}
  .barra{height:7px;display:block}

  .acoes{position:sticky;bottom:0;background:var(--papel);border-top:2px solid var(--tinta);
    padding:14px 0;margin-top:44px;display:flex;gap:11px;flex-wrap:wrap;align-items:center}
  button{font:600 13.5px inherit;padding:10px 17px;border:1px solid var(--tinta);
    background:var(--tinta);color:var(--papel);cursor:pointer;border-radius:0}
  button.sec{background:transparent;color:var(--tinta)}
  button:hover{opacity:.85}
  button:focus-visible{outline:2px solid var(--aqua);outline-offset:2px}
  .aviso{font-size:12.5px;color:var(--tinta-3);flex:1;min-width:200px}

  .ressalva{border:1px solid var(--ambar);background:#FDF6E3;padding:15px 18px;
    margin-top:22px;font-size:13.5px}
  .ressalva b{display:block;margin-bottom:6px;font-size:11px;letter-spacing:.13em;
    text-transform:uppercase;color:var(--ambar)}

  @media print{
    .acoes,button{display:none}
    body{background:#fff}
    textarea{background:#fff;border:1px solid #999}
  }
  @media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
</style>
</head>
<body>
<div class="wrap">

<header>
  <div class="eyebrow">Embratel e Claro Governo · Backlog SGP</div>
  <h1>BACKLOG EM FOCO</h1>
  <div class="meta">
    <span>Report de <b id="m-data"></b></span>
    <span>Comparado com <b id="m-base"></b></span>
    <span>Gerado em <b id="m-gerado"></b></span>
  </div>
</header>

<h2>Números do report</h2>
<p class="sub">Variação em relação ao report anterior.  </p>
<div class="kpis" id="kpis"></div>
<div class="kpi-legenda" id="kpi-legenda"></div>

<h2>Forecast</h2>
<p class="sub" style="max-width:none;white-space:nowrap">Distribuição do backlog pelo horizonte previsto de ativação (Data Previsão Ativação Cliente).</p>
<div class="kpis" id="kpis-forecast"></div>
<div class="matriz" id="matriz-backlog"></div>

<h2>Distribuição do aging</h2>
<p class="sub" style="max-width:none;white-space:nowrap">Circuitos por tempo parado na tarefa atual. A faixa de 60+ dias é a que
justifica cobrança individual.</p>
<div class="regua" id="regua"></div>
<div class="legenda" id="legenda"></div>

<h2>Cobrança por coordenador</h2>
<p class="sub" style="max-width:none;white-space:nowrap">Um bloco por coordenação técnica, com o motivo de cada travamento e os casos
mais antigos para cobrança nominal. Escreva no campo amarelo — a cobrança reaparece no
painel da próxima semana, ao lado do número de então.</p>
<div id="frentes"></div>

<h2>Tipos de pendência no total</h2>
<p class="sub" style="max-width:none;white-space:nowrap">Visão transversal: o que trava o backlog inteiro, independente de coordenação.</p>
<table>
  <thead><tr>
    <th>Pendência</th><th class="num">Circuitos</th><th class="num">Projetos</th>
    <th class="num">Aging médio</th><th class="num">60+ dias</th><th class="num">Δ</th>
  </tr></thead>
  <tbody id="t-pendencias"></tbody>
</table>

<h2>Gargalos por tarefa</h2>
<p class="sub" style="max-width:none;white-space:nowrap">Ordenados por impacto: número de circuitos multiplicado pelo aging médio.</p>
<table>
  <thead><tr>
    <th>Tarefa</th><th class="num">Circuitos</th><th class="num">Projetos</th>
    <th class="num">Aging médio</th><th class="num">Máx.</th><th class="num">Δ</th>
  </tr></thead>
  <tbody id="t-gargalos"></tbody>
</table>

<h2>Clientes</h2>
<table>
  <thead><tr>
    <th>Cliente</th><th class="num">Circuitos</th><th class="num">Projetos</th>
    <th>Em gargalo</th><th class="num">Aging médio</th><th class="num">Δ circ.</th>
  </tr></thead>
  <tbody id="t-clientes"></tbody>
</table>

<div class="ressalva" id="ressalva"></div>

<div class="acoes">
  <button onclick="exportar()">Salvar cobranças</button>
  <button class="sec" onclick="document.getElementById('imp').click()">Carregar cobranças</button>
  <button class="sec" onclick="window.print()">Imprimir</button>
  <input type="file" id="imp" accept=".json" hidden onchange="importar(event)">
  <span class="aviso" id="aviso">As cobranças ficam neste navegador até você salvar.
    Clique em <b>Salvar cobranças</b> para baixar o arquivo e guardá-lo no OneDrive ao
    lado da base.</span>
</div>

</div>
<script>
const DADOS = /*__DADOS__*/;
const S = DADOS.snap, D = DADOS.deltas;
const notas = Object.assign({}, DADOS.anotacoes || {});

const n = v => (v === null || v === undefined) ? "—" : v.toLocaleString("pt-BR");
const esc = s => String(s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const receitaCardMap = new Map([
  ["circuitos", "receita_total"],
  ["gross", "receita_gross"],
  ["servicos", "receita_servicos"],
  ["outros", "receita_outros"],
  ["internet", "receita_internet"],
  ["dados", "receita_dados"],
  ["voz", "receita_voz"],
  ["wifi", "receita_wifi"],
  ["pendentes", "receita_pendentes"],
  ["entregues", "receita_entregues"],
  ["pend_cliente", "receita_pend_cliente"],
  ["pend_vendas", "receita_pend_vendas"],
  ["pend_pje", "receita_pend_pje"],
]);
const receitaFooter = k => {
  const key = receitaCardMap.get(k);
  const value = key ? K[key] : null;
  if (value === null || value === undefined || Number.isNaN(Number(value))) {
    return "";
  }
  return `<div class="kpi-footer kpi-footer-receita">R$ ${Number(value).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>`;
};
const extraFooter = k => {
  if (k !== "clientes") {
    return "";
  }
  return `<div class="kpi-footer kpi-footer-right">Projetos: ${n(K.projetos)}</div>`;
};
const alertaFooter = k => {
  if (k !== "sem_carimbo") {
    return "";
  }
  return `<div class="kpi-alerta"> ⚠️ ${n(K.sem_carimbo_mais_3_dias)} há mais de 3 dias</div>`;
};
const percentualKeys = new Set([
  "gross",
  "servicos",
  "outros",
  "internet",
  "dados",
  "voz",
  "wifi",
  "pendentes",
  "entregues",
  "com_gargalo",
  "fora_sla",
  "pend_cliente",
  "pend_vendas",
  "pend_pje",
  "d60_mais",
  "sem_carimbo",
  "sem_aging",
]);
const pctTotalCard = (k, grp) => {
  if (!percentualKeys.has(k)) {
    return "";
  }
  const total = Number(K.circuitos);
  const valor = Number(K[k]);
  if (!Number.isFinite(total) || total <= 0 || !Number.isFinite(valor)) {
    return "";
  }
  const pct = (valor / total * 100).toLocaleString("pt-BR", {minimumFractionDigits: 1, maximumFractionDigits: 1});
  return `<div class="kpi-pct grp-${grp}">${pct}%</div>`;
};
const cardValorPrincipal = (k, v) => {
  if (k !== "datas_nao_preenchidas") {
    return `<div class="val">${v}</div>`;
  }
  return `<div class="val val-composto"><div class="kpi-split">
    <div class="kpi-split-item">
      <span class="qtd">${n(K.data_prevista_ativacao_a_definir)}</span>
      <span class="txt">Prevista ativação</span>
    </div>
    <div class="kpi-split-item">
      <span class="qtd">${n(K.data_compromisso_sem_data)}</span>
      <span class="txt">Compromisso cliente</span>
    </div>
  </div></div>`;
};
/* Um delta pode ser bom ou ruim dependendo da métrica. `subirEhRuim` diz qual. */
function chipDelta(v, subirEhRuim = true){
  if (v === null || v === undefined) return '<span class="none">sem base</span>';
  if (v === 0) return '<span class="delta flat">sem mudança</span>';
  const piorou = subirEhRuim ? v > 0 : v < 0;
  const sinal = v > 0 ? "+" : "";
  return `<span class="delta ${piorou ? "up" : "down"}">${sinal}${v.toLocaleString("pt-BR")}</span>`;
}

function chipDeltaCard(k, subirEhRuim = true){
  if (!D) return '<span class="none">sem base</span>';
  if (["circuitos", "gross", "servicos", "outros", "internet", "dados", "voz", "wifi"].includes(k) && D.kpis_entradas) {
    return chipDelta(D.kpis_entradas[k], subirEhRuim);
  }
  return chipDelta(D.kpis[k], subirEhRuim);
}

document.getElementById("m-data").textContent = S.data_report;
document.getElementById("m-gerado").textContent = S.gerado_em;
document.getElementById("m-base").textContent = D ? D.data_base : "primeiro report — sem base";

/* KPIs -------------------------------------------------------------------- */
const K = S.kpis;
const pct = (a, b) => b ? (a / b * 100).toFixed(1) + "%" : "—";
const cards = [
  ["Circuitos", n(K.circuitos), "circuitos", true, "base"],
  ["Clientes", n(K.clientes), "clientes", false, "base"],
  ["GROSS", n(K.gross), "gross", false, "classificacao"],
  ["SERVIÇOS", n(K.servicos), "servicos", false, "classificacao"],
  ["OUTROS", n(K.outros), "outros", false, "classificacao"],
  ["INTERNET", n(K.internet), "internet", false, "produto"],
  ["DADOS", n(K.dados), "dados", false, "produto"],
  ["VOZ", n(K.voz), "voz", false, "produto"],
  ["WIFI", n(K.wifi), "wifi", false, "produto"],
  ["Acessos Pendentes", n(K.pendentes), "pendentes", true, "status"],
  ["Acessos Entregues", n(K.entregues), "entregues", false, "status"],
  ["Pendência Cliente", n(K.pend_cliente), "pend_cliente", true, "pendencias"],
  ["Pendência Vendas", n(K.pend_vendas), "pend_vendas", true, "pendencias"],
  ["Pendência PJE", n(K.pend_pje), "pend_pje", true, "pendencias"],
  ["Sem carimbo", n(K.sem_carimbo), "sem_carimbo", true, "dados-ausentes"],
  ["Datas Pendentes", n(K.datas_nao_preenchidas), "datas_nao_preenchidas", true, "dados-ausentes"],
  ["Com gargalo", n(K.com_gargalo), "com_gargalo", true, "risco"],
];
document.getElementById("kpis").innerHTML = cards.map(([l, v, k, ruim, grp]) => `
  <div class="kpi grp-${grp} card-${k}">${pctTotalCard(k, grp)}<div class="lbl">${l}</div>${cardValorPrincipal(k, v)}
  ${k === "datas_nao_preenchidas" ? "" : chipDeltaCard(k, ruim)}${receitaFooter(k)}${extraFooter(k)}${alertaFooter(k)}</div>`).join("");

document.getElementById("kpi-legenda").innerHTML = [
  ["#0E8F8B", "Base"],
  ["#B87410", "Classificação"],
  ["#2D6FA6", "Produto"],
  ["#6E9E5C", "Status"],
  ["#B03F26", "Pendências"],
  ["#7A8791", "Dados Ausentes"],
  ["#B03F26", "Risco"],
].map(([cor, nome]) => `<span><i style="background:${cor}"></i>${nome}</span>`).join("");

/* Forecast (FAIXA_BACKLOG) ------------------------------------------------- */
const totalCircuitosForecast = (S.forecast || []).reduce((sum, item) => sum + (item.circuitos || 0), 0) || K.circuitos;
document.getElementById("kpis-forecast").innerHTML = (S.forecast || []).map(item => {
  const pctForecast = totalCircuitosForecast ? (item.circuitos / totalCircuitosForecast * 100).toLocaleString("pt-BR", {minimumFractionDigits: 1, maximumFractionDigits: 1}) : "0,0";
  const deltaForecast = D && D.forecast && D.forecast[item.id] ? D.forecast[item.id].circuitos : null;
  const receitaForecastHtml = item.receita !== undefined
    ? `<div class="kpi-footer kpi-footer-receita">R$ ${Number(item.receita).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>`
    : "";
  return `<div class="kpi grp-forecast">
    <div class="kpi-pct grp-forecast">${pctForecast}%</div>
    <div class="lbl">${esc(item.tipo)}</div>
    <div class="val">${n(item.circuitos)}</div>
    ${D ? chipDelta(deltaForecast, true) : ""}${receitaForecastHtml}
  </div>`;
}).join("");

const totalBacklog = S.frentes.reduce((sum, f) => sum + (f.circuitos || 0), 0);
const totalEntradasBacklog = D && D.entradas !== undefined ? D.entradas : null;
const hasEntradasBacklog = totalEntradasBacklog !== null;
const totalReceita = S.frentes.reduce((sum, f) => sum + (f.receita || 0), 0);
const hasReceita = S.frentes.some(f => f.receita !== undefined);
const totalGross = S.frentes.reduce((sum, f) => sum + (f.gross || 0), 0);
const totalServicos = S.frentes.reduce((sum, f) => sum + (f.servicos || 0), 0);
const totalOutros = S.frentes.reduce((sum, f) => sum + (f.outros || 0), 0);
const totalInternet = S.frentes.reduce((sum, f) => sum + (f.internet || 0), 0);
const totalDados = S.frentes.reduce((sum, f) => sum + (f.dados || 0), 0);
const totalVoz = S.frentes.reduce((sum, f) => sum + (f.voz || 0), 0);
const totalWifi = S.frentes.reduce((sum, f) => sum + (f.wifi || 0), 0);
const totalAcessosPendentes = S.frentes.reduce((sum, f) => sum + (f.pendentes || 0), 0);
const totalAcessosEntregues = S.frentes.reduce((sum, f) => sum + (f.entregues || 0), 0);
const totalPendCliente = S.frentes.reduce((sum, f) => sum + (f.pend_cliente || 0), 0);
const totalPendVendas = S.frentes.reduce((sum, f) => sum + (f.pend_vendas || 0), 0);
const totalPendPje = S.frentes.reduce((sum, f) => sum + (f.pend_pje || 0), 0);
const totalMatrizSemCarimbo = S.frentes.reduce((sum, f) => sum + (f.sem_carimbo || 0), 0);
const totalMatrizPrevistaAtivacao = S.frentes.reduce((sum, f) => sum + (f.data_prevista_ativacao || 0), 0);
const totalMatrizCompromissoCliente = S.frentes.reduce((sum, f) => sum + (f.data_compromisso_cliente || 0), 0);
const maxMatriz = key => Math.max(...S.frentes.map(f => Number(f[key]) || 0), 1);
const maxMatrizForecast = id => Math.max(...S.frentes.map(f => Number(f.forecast && f.forecast[id]) || 0), 1);
const matrizBar = (valor, max, grupo) => {
  if (valor === null || valor === undefined) {
    return '<span class="none">sem base</span>';
  }
  const largura = Math.max(0, Math.min(100, Number(valor) / max * 100));
  return `<span class="matriz-bar ${grupo}" style="--bar-width:${largura}%"><span>${n(valor)}</span></span>`;
};

document.getElementById("matriz-backlog").innerHTML = `
  <h3>Backlog por Gerente Técnico (Valores Atuais e Variações vs Base)</h3>
  <table class="matriz-table">
    <colgroup>
      <col style="width:168px">
      <col style="width:65px">
      <col style="width:78px">
      <col style="width:126px">
      <col style="width:64px">
      <col style="width:72px">
      <col style="width:64px">
      <col style="width:72px">
      <col style="width:60px">
      <col style="width:56px">
      <col style="width:56px">
      <col style="width:84px">
      <col style="width:84px">
      <col style="width:66px">
      <col style="width:66px">
      <col style="width:58px">
      <col style="width:88px">
      <col style="width:100px">
      <col style="width:105px">
      ${(S.forecast || []).map(() => '<col style="width:84px">').join("")}
    </colgroup>
    <thead>
      <tr class="group-label">
        <th rowspan="2">Gerente Técnico</th>
        <th class="group backlog" colspan="3">BACKLOG TOTAL</th>
        <th class="group classificacao" colspan="3">CLASSIFICAÇÃO</th>
        <th class="group produto" colspan="4">PRODUTO</th>
        <th class="group status" colspan="2">STATUS ACESSOS</th>
        <th class="group pendencias" colspan="3">PENDÊNCIAS</th>
        <th class="group dados-ausentes" colspan="3">DADOS AUSENTES</th>
        <th class="group forecast" colspan="${(S.forecast || []).length}">FORECAST</th>
      </tr>
      <tr class="group-header"><th class="num backlog group-header-total">Total</th><th class="num backlog">Entradas</th><th class="money backlog">Receita</th><th class="num classificacao">GROSS</th><th class="num classificacao">SERVIÇOS</th><th class="num classificacao">OUTROS</th><th class="num produto">INTERNET</th><th class="num produto">DADOS</th><th class="num produto">VOZ</th><th class="num produto">WIFI</th><th class="num status">Pendentes</th><th class="num status">Entregues</th><th class="num pendencias">Cliente</th><th class="num pendencias">Vendas</th><th class="num pendencias">PJE</th><th class="num dados-ausentes">Sem carimbo</th><th class="num dados-ausentes">Prevista ativação</th><th class="num dados-ausentes">Compromisso cliente</th>${(S.forecast || []).map((item, idx) => `<th class="num forecast${idx === 0 ? " forecast-start" : ""}">${esc(item.tipo)}</th>`).join("")}</tr>
    </thead>
    <tbody>
      ${S.frentes.map(f => {
        const d = D && D.frentes[f.id];
        const entradasCircuitos = d && d.entradas !== undefined ? d.entradas : null;
        const qtdDeltaHtml = entradasCircuitos !== null
          ? chipDelta(entradasCircuitos, true)
          : '<span class="none">sem base</span>';
        const receitaAbsHtml = f.receita !== undefined
          ? 'R$ ' + Number(f.receita).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2})
          : '<span class="none">sem base</span>';
        return `
        <tr>
          <td>${esc(f.coordenador)}</td>
          <td class="num">${matrizBar(f.circuitos, maxMatriz("circuitos"), "backlog")}</td>
          <td class="num">${qtdDeltaHtml}</td>
          <td class="money">${receitaAbsHtml}</td>
          <td class="num">${matrizBar(f.gross, maxMatriz("gross"), "classificacao")}</td>
          <td class="num">${matrizBar(f.servicos, maxMatriz("servicos"), "classificacao")}</td>
          <td class="num">${matrizBar(f.outros, maxMatriz("outros"), "classificacao")}</td>
          <td class="num">${matrizBar(f.internet, maxMatriz("internet"), "produto")}</td>
          <td class="num">${matrizBar(f.dados, maxMatriz("dados"), "produto")}</td>
          <td class="num">${matrizBar(f.voz, maxMatriz("voz"), "produto")}</td>
          <td class="num">${matrizBar(f.wifi, maxMatriz("wifi"), "produto")}</td>
          <td class="num">${matrizBar(f.pendentes, maxMatriz("pendentes"), "status")}</td>
          <td class="num">${matrizBar(f.entregues, maxMatriz("entregues"), "status")}</td>
          <td class="num">${matrizBar(f.pend_cliente, maxMatriz("pend_cliente"), "pendencias")}</td>
          <td class="num">${matrizBar(f.pend_vendas, maxMatriz("pend_vendas"), "pendencias")}</td>
          <td class="num">${matrizBar(f.pend_pje, maxMatriz("pend_pje"), "pendencias")}</td>
          <td class="num">${matrizBar(f.sem_carimbo, maxMatriz("sem_carimbo"), "dados-ausentes")}</td>
          <td class="num">${matrizBar(f.data_prevista_ativacao, maxMatriz("data_prevista_ativacao"), "dados-ausentes")}</td>
          <td class="num">${matrizBar(f.data_compromisso_cliente, maxMatriz("data_compromisso_cliente"), "dados-ausentes")}</td>
          ${(S.forecast || []).map((item, idx) => `<td class="num${idx === 0 ? " forecast-start" : ""}">${matrizBar(f.forecast ? f.forecast[item.id] : undefined, maxMatrizForecast(item.id), "forecast")}</td>`).join("")}
        </tr>`;
      }).join("")}
      <tr class="total">
        <td>Total</td>
        <td class="num">${matrizBar(totalBacklog, maxMatriz("circuitos"), "backlog")}</td>
        <td class="num">${hasEntradasBacklog ? chipDelta(totalEntradasBacklog, true) : '<span class="none">sem base</span>'}</td>
        <td class="money">${hasReceita ? 'R$ ' + Number(totalReceita).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '<span class="none">sem base</span>'}</td>
        <td class="num">${matrizBar(totalGross, maxMatriz("gross"), "classificacao")}</td>
        <td class="num">${matrizBar(totalServicos, maxMatriz("servicos"), "classificacao")}</td>
        <td class="num">${matrizBar(totalOutros, maxMatriz("outros"), "classificacao")}</td>
        <td class="num">${matrizBar(totalInternet, maxMatriz("internet"), "produto")}</td>
        <td class="num">${matrizBar(totalDados, maxMatriz("dados"), "produto")}</td>
        <td class="num">${matrizBar(totalVoz, maxMatriz("voz"), "produto")}</td>
        <td class="num">${matrizBar(totalWifi, maxMatriz("wifi"), "produto")}</td>
        <td class="num">${matrizBar(totalAcessosPendentes, maxMatriz("pendentes"), "status")}</td>
        <td class="num">${matrizBar(totalAcessosEntregues, maxMatriz("entregues"), "status")}</td>
        <td class="num">${matrizBar(totalPendCliente, maxMatriz("pend_cliente"), "pendencias")}</td>
        <td class="num">${matrizBar(totalPendVendas, maxMatriz("pend_vendas"), "pendencias")}</td>
        <td class="num">${matrizBar(totalPendPje, maxMatriz("pend_pje"), "pendencias")}</td>
        <td class="num">${matrizBar(totalMatrizSemCarimbo, maxMatriz("sem_carimbo"), "dados-ausentes")}</td>
        <td class="num">${matrizBar(totalMatrizPrevistaAtivacao, maxMatriz("data_prevista_ativacao"), "dados-ausentes")}</td>
        <td class="num">${matrizBar(totalMatrizCompromissoCliente, maxMatriz("data_compromisso_cliente"), "dados-ausentes")}</td>
        ${(S.forecast || []).map((item, idx) => `<td class="num${idx === 0 ? " forecast-start" : ""}">${matrizBar(item.circuitos, maxMatrizForecast(item.id), "forecast")}</td>`).join("")}
      </tr>
    </tbody>
  </table>`;

/* Régua de aging ---------------------------------------------------------- */
const faixas = [
  ["até 15 dias", K.ate_15, "#0E8F8B"],
  ["16–30 dias", K.d16_30, "#6E9E5C"],
  ["31–60 dias", K.d31_60, "#C4890A"],
  ["60+ dias", K.d60_mais, "#B03F26"],
];
const totalF = faixas.reduce((a, f) => a + f[1], 0);
document.getElementById("regua").innerHTML = faixas.map(([l, v, c]) =>
  `<div style="background:${c};flex:${v}" title="${l}: ${n(v)}">${v / totalF > .08 ? n(v) : ""}</div>`).join("");
document.getElementById("legenda").innerHTML = faixas.map(([l, v, c]) =>
  `<span><i class="sw" style="background:${c}"></i>${l} · <b>${n(v)}</b> (${pct(v, totalF)})</span>`).join("");

/* Frentes de cobrança — uma por coordenador ------------------------------- */
document.getElementById("frentes").innerHTML = S.frentes.map(f => {
  const d = D && D.frentes[f.id] || {};
  const grau = f.pct_gargalo >= 45 ? "crit" : f.pct_gargalo >= 25 ? "aten" : "";
  const antiga = notas[f.id];
  const iniciais = f.coordenador.split(" ").map(p => p[0]).join("").slice(0, 2).toUpperCase();

  const pend = f.pendencias.map(p => `
    <div class="pend">
      <div class="pend-t">${esc(p.tipo)}</div>
      <div class="pend-n">${n(p.circuitos)}</div>
      <div class="pend-a">${p.aging_medio !== null ? p.aging_medio + " d" : "—"}</div>
    </div>`).join("");

  const piores = f.piores.map(c => `
    <li><span class="mono">${esc(c.designacao)}</span>
      <b>${n(c.aging)} d</b> · ${esc(c.cliente)}
      ${c.carimbo ? '· carimbo <span class="mono">' + esc(c.carimbo) + "</span>" : ""}
      <span class="tar">${esc(c.tarefa)}</span></li>`).join("");

  return `<div class="frente ${grau}">
    <div class="fhead">
      <div class="fid">
        <span class="avatar">${esc(iniciais)}</span>
        <div><div class="fnome">${esc(f.coordenador)}</div>
          <div class="fcoord">Coordenação técnica</div></div>
      </div>
      <div class="fnums">
        <div class="fnum"><div class="n">${n(f.circuitos)}</div><div class="l">circuitos</div>
          ${D ? chipDelta(d.circuitos, true) : ""}</div>
        <div class="fnum"><div class="n">${n(f.projetos)}</div><div class="l">projetos</div></div>
        <div class="fnum"><div class="n">${f.pct_gargalo}<small>%</small></div>
          <div class="l">em gargalo</div>${D ? chipDelta(d.pct_gargalo, true) : ""}</div>
        <div class="fnum"><div class="n">${f.aging_medio}</div><div class="l">aging médio</div>
          ${D ? chipDelta(d.aging_medio, true) : ""}</div>
        <div class="fnum"><div class="n">${n(f.criticos)}</div><div class="l">60+ dias</div>
          ${D ? chipDelta(d.criticos, true) : ""}</div>
      </div>
    </div>

    <div class="secao">
      <div class="secao-t">Por que está parado</div>
      <div class="pends">${pend}</div>
    </div>

    ${piores ? `<div class="secao">
      <div class="secao-t">Casos mais antigos</div>
      <ul class="piores">${piores}</ul></div>` : ""}

    <div class="cobranca">
      <label for="nota-${esc(f.id)}">Cobrança desta semana</label>
      <textarea id="nota-${esc(f.id)}" data-id="${esc(f.id)}"
        placeholder="O que foi cobrado e até quando.">${esc(antiga ? antiga.texto : "")}</textarea>
      ${antiga ? `<div class="hist">Registrado em <b>${esc(antiga.data)}</b></div>` : ""}
    </div>
  </div>`;
}).join("");

/* Pendências no total ----------------------------------------------------- */
document.getElementById("t-pendencias").innerHTML = S.pendencias_geral.map(p => {
  const d = D && D.pendencias_geral[p.id] || {};
  return `<tr><td>${esc(p.tipo)}</td>
    <td class="num">${n(p.circuitos)}</td><td class="num">${n(p.projetos)}</td>
    <td class="num">${p.aging_medio ?? "—"}</td><td class="num">${n(p.criticos)}</td>
    <td class="num">${D ? chipDelta(d.circuitos, true) : "—"}</td></tr>`;
}).join("");

/* Tabelas ----------------------------------------------------------------- */
document.getElementById("t-gargalos").innerHTML = S.gargalos.map(g => {
  const d = D && D.gargalos[g.id] || {};
  return `<tr>
    <td><span class="mono">${esc(g.id)}</span> ${esc(g.tarefa.replace(g.id, "").replace(/^\s*-\s*/, ""))}</td>
    <td class="num">${n(g.circuitos)}</td><td class="num">${n(g.projetos)}</td>
    <td class="num">${g.aging_medio}</td><td class="num">${n(g.aging_max)}</td>
    <td class="num">${D ? chipDelta(d.circuitos, true) : "—"}</td></tr>`;
}).join("");

document.getElementById("t-clientes").innerHTML = S.clientes.map(c => {
  const d = D && D.clientes[c.id] || {};
  const cor = c.pct_gargalo >= 50 ? "#B03F26" : c.pct_gargalo >= 30 ? "#C4890A" : "#0E8F8B";
  return `<tr>
    <td>${esc(c.cliente)}</td><td class="num">${n(c.circuitos)}</td>
    <td class="num">${n(c.projetos)}</td>
    <td><span class="barra-bg"><span class="barra" style="width:${c.pct_gargalo}%;background:${cor}"></span></span>${c.pct_gargalo}%</td>
    <td class="num">${c.aging_medio ?? "—"}</td>
    <td class="num">${D ? chipDelta(d.circuitos, true) : "—"}</td></tr>`;
}).join("");

/* Ressalvas --------------------------------------------------------------- */
const semAging = (K.sem_aging / K.circuitos * 100).toFixed(1);
document.getElementById("ressalva").innerHTML =
  `<b>Ressalvas sobre os dados</b>
   ${n(K.sem_aging)} circuitos (${semAging}% da base) estão sem aging da tarefa atual e
   ficam de fora de todo cálculo de aging e SLA — as métricas acima consideram
   ${n(K.circuitos - K.sem_aging)} circuitos.
   ${D ? `Entre ${D.data_base} e ${S.data_report}: <b>${n(D.entradas)}</b> circuitos entraram
   e <b>${n(D.saidas)}</b> saíram do backlog.`
   : "Este é o primeiro report gravado — não há comparação disponível. A partir da próxima execução o painel mostra as variações."}`;

/* Cobranças --------------------------------------------------------------- */
document.addEventListener("input", e => {
  if (e.target.tagName !== "TEXTAREA") return;
  const id = e.target.dataset.id, txt = e.target.value.trim();
  if (txt) notas[id] = {texto: txt, data: S.data_report, autor: notas[id]?.autor || ""};
  else delete notas[id];
  document.getElementById("aviso").innerHTML =
    "Cobranças alteradas. Clique em <b>Salvar cobranças</b> para não perder.";
});

function exportar(){
  const blob = new Blob([JSON.stringify(notas, null, 2)], {type: "application/json"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "anotacoes.json";
  a.click();
  URL.revokeObjectURL(a.href);
  document.getElementById("aviso").innerHTML =
    "Arquivo baixado. Guarde <b>anotacoes.json</b> na mesma pasta do gerador.";
}

function importar(ev){
  const f = ev.target.files[0];
  if (!f) return;
  const r = new FileReader();
  r.onload = () => {
    try {
      const carregado = JSON.parse(r.result);
      Object.assign(notas, carregado);
      document.querySelectorAll("textarea[data-id]").forEach(t => {
        if (carregado[t.dataset.id]) t.value = carregado[t.dataset.id].texto;
      });
      document.getElementById("aviso").textContent = "Cobranças carregadas.";
    } catch {
      document.getElementById("aviso").textContent =
        "Não foi possível ler o arquivo. Use o anotacoes.json gerado por este painel.";
    }
  };
  r.readAsText(f);
}
</script>
</body>
</html>
"""


# --------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        raise SystemExit(
            "Uso: python gera_painel_backlog.py <base.xlsx> [pasta_destino_opcional]"
        )

    base = Path(sys.argv[1])
    if not base.exists():
        raise SystemExit(f"ERRO: arquivo nao encontrado: {base}")

    print(f"Lendo {base} ...")
    df = carregar_base(base)
    print(f"  {len(df)} circuitos, {df[COL_PROJ].dropna().nunique()} projetos (contagem por CARIMBO_PREFIXO)")

    snap = calcular_metricas(df)
    anterior = snapshot_anterior(snap["data_report"])
    deltas = calcular_deltas(snap, anterior)

    caminho_snap = salvar_snapshot(snap)
    print(f"Snapshot gravado: {caminho_snap}")

    if deltas:
        print(f"Comparando com {deltas['data_base']}: "
              f"+{deltas['entradas']} entradas, -{deltas['saidas']} saidas")
    else:
        print("Sem snapshot anterior — este painel sai sem comparacao.")

    saida = gerar_html(snap, deltas, carregar_anotacoes())
    print(f"Painel gerado: {saida}")
    print(f"Pasta de destino usada nesta execução: {BASE_DIR.resolve()}")


if __name__ == "__main__":
    main()




##========================================================================================

##Executar no CMD:
# python gera_painel_backlog.py "C:\Users\Z181040\OneDrive - Claro SA\BASES\Projeto_Cloude\Relatorios_Semanais_Backlog\BD_Backlog_SGP.xlsx"

# cd C:\Users\Z181040\OneDrive - Claro SA\BASES\Projeto_Cloude\Relatorios_Semanais_Backlog
# python gera_painel_backlog.py BD_Backlog_SGP.xlsx
