import hashlib
import io
import random
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

INVALIDOS = {"NAN", "NONE", "NULL", ""}
PADRAO_CANCELADO = re.compile(r"NAO VOU|NAO VAI|CANCELAD")
PADRAO_2130 = re.compile(r"21\s*[:H.]\s*30", re.IGNORECASE)

REGRAS_POLO = [
    ("ANHANGUERA-SEDE", re.compile(r"ANHANGUERA.*SEDE|SEDE.*ANHANGUERA")),
    ("ANHANGUERA-CAP", re.compile(r"ANHANGUERA.*CAP|CAP.*ANHANGUERA")),
    ("ANHANGUERA-SEDE", re.compile(r"ANHANGUERA")),
    ("EXPERT/EVOLUA", re.compile(r"EXPERT|EVOLUA")),
    ("HUMBERTO RIBEIRO", re.compile(r"HUMBERTO")),
    ("UESB", re.compile(r"UESB")),
    ("M1", re.compile(r"UNEX ?M1|\bM1\b|SAO JUDAS")),
    ("M3", re.compile(r"UNEX ?M3|\bM3\b|CESAR BORGES")),
    ("UNEX MEDICINA", re.compile(r"MEDICINA")),
    ("UNIGRANDE", re.compile(r"UNIGRANDE")),
    ("UFSB", re.compile(r"UFSB")),
    ("IFBA", re.compile(r"IFBA")),
    ("FATESB", re.compile(r"FATESB")),
    ("PROCURSO", re.compile(r"PROCURSO")),
]

MATRIZ_MISTURA = {
    "M1": ["UNEX MEDICINA", "FATESB", "UNIGRANDE", "EXPERT/EVOLUA", "PROCURSO", "M3"],
    "UNEX MEDICINA": ["M1", "FATESB", "UNIGRANDE", "EXPERT/EVOLUA", "PROCURSO", "M3", "ANHANGUERA-SEDE"],
    "M3": ["PROCURSO", "EXPERT/EVOLUA", "FATESB", "UNIGRANDE", "UESB", "M1", "ANHANGUERA-SEDE"],
    "ANHANGUERA-SEDE": ["ANHANGUERA-CAP", "UFSB", "IFBA", "EXPERT/EVOLUA", "PROCURSO", "HUMBERTO RIBEIRO", "UESB", "M3"],
    "ANHANGUERA-CAP": ["ANHANGUERA-SEDE", "UFSB", "IFBA", "EXPERT/EVOLUA", "PROCURSO", "HUMBERTO RIBEIRO", "UESB"],
    "EXPERT/EVOLUA": ["PROCURSO", "FATESB", "UNIGRANDE", "M3", "HUMBERTO RIBEIRO", "UESB", "M1", "UNEX MEDICINA", "ANHANGUERA-SEDE"],
    "PROCURSO": ["EXPERT/EVOLUA", "FATESB", "UNIGRANDE", "M3", "HUMBERTO RIBEIRO", "UESB", "M1", "UNEX MEDICINA", "ANHANGUERA-SEDE"],
    "UNIGRANDE": ["FATESB", "EXPERT/EVOLUA", "PROCURSO", "M3", "M1", "HUMBERTO RIBEIRO", "UESB"],
    "FATESB": ["UNIGRANDE", "EXPERT/EVOLUA", "PROCURSO", "M3", "M1", "UNEX MEDICINA", "HUMBERTO RIBEIRO", "UESB", "ANHANGUERA-SEDE"],
    "UESB": ["HUMBERTO RIBEIRO", "M3", "EXPERT/EVOLUA", "PROCURSO", "FATESB", "UNIGRANDE", "ANHANGUERA-SEDE", "ANHANGUERA-CAP"],
    "HUMBERTO RIBEIRO": ["UESB", "M3", "EXPERT/EVOLUA", "PROCURSO", "UNIGRANDE", "ANHANGUERA-SEDE", "ANHANGUERA-CAP"],
    "IFBA": ["ANHANGUERA-SEDE", "ANHANGUERA-CAP", "UFSB", "UNIGRANDE", "FATESB", "EXPERT/EVOLUA", "PROCURSO", "M3", "HUMBERTO RIBEIRO", "UESB"],
    "UFSB": ["ANHANGUERA-SEDE", "ANHANGUERA-CAP", "IFBA", "EXPERT/EVOLUA", "FATESB"],
    "OUTROS": [],
}

MEDICINA_PODE_SER_ANCORA = False

_PREF_ONIBUS = ["M1", "M3", "ANHANGUERA-SEDE", "ANHANGUERA-CAP", "UNIGRANDE", "PROCURSO", "UNEX MEDICINA", "OUTROS"]
_PREF_MICROS = ["M1", "M3", "ANHANGUERA-SEDE", "UESB", "ANHANGUERA-CAP", "UNIGRANDE", "HUMBERTO RIBEIRO",
                "UNEX MEDICINA", "EXPERT/EVOLUA", "PROCURSO", "FATESB", "IFBA", "UFSB", "OUTROS"]
PREF_ONIBUS = [p for p in _PREF_ONIBUS if MEDICINA_PODE_SER_ANCORA or p != "UNEX MEDICINA"]
PREF_MICROS = [p for p in _PREF_MICROS if MEDICINA_PODE_SER_ANCORA or p != "UNEX MEDICINA"]

PADRAO_MISTA = re.compile(
    r"\bIDA\b\s*[:\-]?\s*(.+?)(?:\s+E\s+|\s*[/,;+\-]\s*)(?:RETORNO|VOLTA)\b\s*[:\-]?\s*(.+)",
    re.IGNORECASE,
)
NOMES_CANONICOS = [
    (re.compile(r"\bM1\b|SAO JUDAS"), "UNEX M1 SÃO JUDAS"),
    (re.compile(r"\bM3\b|CESAR BORGES"), "UNEX M3 CÉSAR BORGES"),
    (re.compile(r"MEDICINA"), "UNEX MEDICINA"),
]

PALETA_CORES = [
    "FF66FF", "92D050", "FFC896", "9DC3E6", "C6E0B4",
    "F4B183", "B4A7D6", "8EA9DB", "F8CBAD", "BFBFBF",
    "E6B8AF", "C5E1A5", "FFE082", "B39DDB", "80CBC4",
]
COR_PREDEFINIDA = {
    "UNIGRANDE": "FF66FF", "ANHANGUERA": "92D050", "M1": "FFC896",
    "M3": "9DC3E6", "UESB": "C6E0B4", "EXPERT": "F4B183",
    "HUMBERTO": "B4A7D6", "MEDICINA": "80CBC4", "OUTROS": "BFBFBF",
}

SO_UM_SENTIDO = {"SÓ RETORNO", "SÓ IDA"}
INDICE_GRUPO = {"SÓ IDA": 0, "SÓ RETORNO": 1}


def normalizar_chave(txt):
    t = unicodedata.normalize("NFKD", str(txt))
    t = "".join(c for c in t if not unicodedata.combining(c)).upper()
    return " ".join(re.sub(r"[^A-Z0-9]+", " ", t).split())


CHAVE_CANONICA = {normalizar_chave(n): n for _, n in NOMES_CANONICOS}


def limpar_instituicao(txt):
    t = unicodedata.normalize("NFC", str(txt))
    t = re.sub(r"[^\w\s\-/().,ºª&]", " ", t)
    return " ".join(t.split()).upper()


def padronizar_instituicao(txt):
    t = limpar_instituicao(txt)
    return CHAVE_CANONICA.get(normalizar_chave(t), t)


def identificar_polo(inst):
    chave = normalizar_chave(inst)
    for polo, padrao in REGRAS_POLO:
        if padrao.search(chave):
            return polo
    return "OUTROS"


def classificar_uso(uso):
    u = normalizar_chave(uso)
    tem_ida = bool(re.search(r"\bIDA\b", u))
    tem_volta = "RETORNO" in u or "VOLTA" in u
    if tem_ida and tem_volta:
        return "IDA E RETORNO"
    if tem_volta:
        return "SÓ RETORNO"
    if tem_ida:
        return "SÓ IDA"
    return "OUTRO"


def nome_canonico(lado):
    chave = normalizar_chave(lado)
    for padrao, canonico in NOMES_CANONICOS:
        if padrao.search(chave):
            return canonico
    return " ".join(lado.split()).upper()


def dividir_instituicao_mista(inst, uso):
    m = PADRAO_MISTA.search(inst)
    if not m:
        return [(inst, uso)]
    return [(nome_canonico(m.group(1)), "SÓ IDA"), (nome_canonico(m.group(2)), "SÓ RETORNO")]


def bairro_exibicao(aluno):
    return "NÃO RETORNO" if aluno["grupo"] == "SÓ IDA" else aluno["bairro"]


def contar_ida_volta(alunos):
    c = Counter(a["grupo"] for a in alunos)
    return c["IDA E RETORNO"] + c["SÓ IDA"], c["IDA E RETORNO"] + c["SÓ RETORNO"]


def preparar_alunos(df, col):
    df = df.dropna(how="all").reset_index(drop=True)
    alunos, descartados, uid = [], [], 0

    for ordem, linha in enumerate(df.to_dict("records")):
        def txt(chave):
            v = linha.get(col[chave])
            return "" if pd.isna(v) else str(v).strip()

        nome, inst_txt = txt("nome"), padronizar_instituicao(txt("inst"))
        motivo = None
        if nome.upper() in INVALIDOS:
            motivo = "Nome em branco"
        elif PADRAO_CANCELADO.search(normalizar_chave(nome)):
            motivo = "Marcado como não vai / cancelado"
        elif inst_txt in INVALIDOS:
            motivo = "Instituição em branco"

        if motivo:
            descartados.append({"Linha na planilha": ordem + 2, "Nome": nome,
                                "Instituição (original)": txt("inst"), "Motivo": motivo})
            continue

        bairro = txt("bairro")
        base = {
            "ordem": ordem, "hora": txt("hora"), "nome": nome, "retorno": txt("retorno"),
            "bairro": "" if bairro.upper() in INVALIDOS else bairro, "remanejado": False,
        }
        for inst, uso in dividir_instituicao_mista(inst_txt, txt("uso")):
            uid += 1
            alunos.append({**base, "uid": uid, "inst": inst, "uso": uso,
                           "polo": identificar_polo(inst), "grupo": classificar_uso(uso)})
    return alunos, descartados


class Carga:

    def __init__(self):
        self.por_inst = {}
        self.total = 0

    @staticmethod
    def _assentos(c):
        return max(c[0], c[1]) + c[2]

    def _depois(self, aluno):
        c = list(self.por_inst.get(aluno["inst"], (0, 0, 0)))
        c[INDICE_GRUPO.get(aluno["grupo"], 2)] += 1
        return c

    def custo(self, aluno):
        antes = self.por_inst.get(aluno["inst"], (0, 0, 0))
        return self._assentos(self._depois(aluno)) - self._assentos(antes)

    def adicionar(self, aluno):
        self.total += self.custo(aluno)
        self.por_inst[aluno["inst"]] = self._depois(aluno)


@dataclass
class Veiculo:
    id: str
    tipo: str
    cap: int
    horario: str
    motorista: str = ""
    alunos: list = field(default_factory=list)
    monitor_uid: Optional[int] = None
    carga: Carga = field(default_factory=Carga)
    insts: set = field(default_factory=set)
    polos: set = field(default_factory=set)

    @property
    def ocupacao(self):
        return self.carga.total

    def cabe(self, aluno):
        return self.carga.total + self.carga.custo(aluno) <= self.cap

    def adicionar(self, aluno):
        self.carga.adicionar(aluno)
        self.alunos.append(aluno)
        self.insts.add(aluno["inst"])
        self.polos.add(aluno["polo"])


def selecionar_por_ordem(fila, vagas):
    carga, dentro, fora = Carga(), [], []
    for a in fila:
        if carga.total + carga.custo(a) <= vagas:
            carga.adicionar(a)
            dentro.append(a)
        else:
            fora.append(a)
    return dentro, fora


def afinidade(veiculo, aluno):
    if aluno["inst"] in veiculo.insts:
        return 3
    if aluno["polo"] in veiculo.polos:
        return 2
    if any(aluno["polo"] in MATRIZ_MISTURA.get(p, []) or p in MATRIZ_MISTURA.get(aluno["polo"], [])
           for p in veiculo.polos):
        return 1
    return 0


def completar_vagas(veiculos, restantes):
    sobras = []
    for a in restantes:
        melhor, melhor_chave = None, None
        for v in veiculos:
            if not v.cabe(a):
                continue
            chave = (afinidade(v, a), v.cap - v.ocupacao)
            if melhor is None or chave > melhor_chave:
                melhor, melhor_chave = v, chave
        if melhor is None:
            sobras.append(a)
        else:
            melhor.adicionar(a)
    return sobras


def alocar_frota(fila, frota):
    restantes, veiculos = list(fila), []

    for i, cfg in enumerate(frota):
        if not restantes:
            break
        v = Veiculo(id=f"{cfg['tipo']} {i + 1}", tipo=cfg["tipo"], cap=cfg["cap"],
                    horario=cfg["horario"], motorista=cfg["motorista"])
        veiculos.append(v)

        demanda = Counter(a["polo"] for a in restantes)
        prefs = PREF_ONIBUS if v.tipo == "Ônibus" else PREF_MICROS
        candidatos = sorted((p for p in prefs if demanda[p] > 0),
                            key=lambda p: (demanda[p], -prefs.index(p)), reverse=True)

        ancora = None
        for polo in candidatos:
            if v.tipo == "Ônibus" and polo == "UNIGRANDE" and any(a["polo"] != "UNIGRANDE" for a in restantes):
                continue
            ancora = polo
            break
        if ancora is None:
            ancora = restantes[0]["polo"]

        polos_para_puxar = sorted(dict.fromkeys([ancora] + MATRIZ_MISTURA.get(ancora, [])),
                                  key=lambda p: (demanda[p], p == ancora), reverse=True)
        for polo in polos_para_puxar:
            if demanda[polo] == 0:
                continue
            for a in [x for x in restantes if x["polo"] == polo]:
                if v.cabe(a):
                    v.adicionar(a)

        tomados = {a["uid"] for a in v.alunos}
        restantes = [a for a in restantes if a["uid"] not in tomados]

    sobras = completar_vagas(veiculos, restantes)
    return [v for v in veiculos if v.alunos], sobras


@dataclass
class Resultado:
    fila_2130: list
    fila_22: list
    roteiros_2130: list
    roteiros_22: list
    espera_2130: list
    sem_vaga: list

    @property
    def roteiros(self):
        return self.roteiros_2130 + self.roteiros_22


def gerar_roteiros(alunos, frota_2130, frota_22):
    por_ordem = lambda a: (a["ordem"], a["uid"])
    fila = sorted(alunos, key=por_ordem)

    fila_2130 = [a for a in fila if PADRAO_2130.search(a["retorno"])]
    fila_normal = [a for a in fila if not PADRAO_2130.search(a["retorno"])]

    movidos = [a for a in fila_2130
               if "MEDICINA" in a["inst"] and not ("M1" in a["inst"] or a["grupo"] == "SÓ IDA")]
    ids_movidos = {a["uid"] for a in movidos}
    fila_2130 = [a for a in fila_2130 if a["uid"] not in ids_movidos]
    fila_normal = sorted(fila_normal + movidos, key=por_ordem)

    contemplados, espera = selecionar_por_ordem(fila_2130, sum(v["cap"] for v in frota_2130))
    roteiros_2130, sobras = alocar_frota(contemplados, frota_2130)
    espera = sorted(espera + sobras, key=por_ordem)
    for a in espera:
        a["remanejado"] = True

    fila_22 = espera + fila_normal
    aceitos, sem_vaga_corte = selecionar_por_ordem(fila_22, sum(v["cap"] for v in frota_22))
    roteiros_22, sobras_22 = alocar_frota(aceitos, frota_22)
    sem_vaga = sorted(sem_vaga_corte + sobras_22, key=por_ordem)

    return Resultado(fila_2130, fila_22, roteiros_2130, roteiros_22, espera, sem_vaga)


def sortear_monitores(roteiros, semente):
    for rot in roteiros:
        if rot.alunos:
            rng = random.Random(f"{semente}-{rot.horario}-{rot.id}")
            rot.monitor_uid = rng.choice(rot.alunos)["uid"]


def construir_mapa_cores(instituicoes):
    mapa, usadas = {}, set(COR_PREDEFINIDA.values())
    for inst in instituicoes:
        chave = normalizar_chave(inst)
        cor = next((c for k, c in COR_PREDEFINIDA.items() if k in chave), None)
        if cor is None:
            cor = next((c for c in PALETA_CORES if c not in usadas), "D9D9D9")
            usadas.add(cor)
        mapa[inst] = cor
    return mapa


def alunos_da_instituicao(rot, inst):
    lista = [a for a in rot.alunos if a["inst"] == inst]
    return [a for a in lista if not a["remanejado"]] + [a for a in lista if a["remanejado"]]


def instituicoes_do_roteiro(rot):
    return list(dict.fromkeys(a["inst"] for a in rot.alunos))


def texto_motorista(rot, preenchimento):
    mot = rot.motorista.strip().upper()
    return mot if mot and mot != "NAN" else preenchimento


def tabela_alunos(alunos, com_horario=False):
    colunas = ["Data/Hora", "Nome do Associado", "Instituição", "Uso do Transporte", "Bairro"]
    linhas = [[a["hora"], a["nome"], a["inst"], a["uso"], a["bairro"]] for a in alunos]
    if com_horario:
        colunas.append("Horário solicitado")
        for linha, a in zip(linhas, alunos):
            linha.append(a["retorno"])
    return pd.DataFrame(linhas, columns=colunas)


def tabela_conferencia(alunos, res):
    def contar(lista):
        return Counter(a["inst"] for a in lista)

    lidos = contar(alunos)
    em_2130 = contar([a for r in res.roteiros_2130 for a in r.alunos])
    em_22 = contar([a for r in res.roteiros_22 for a in r.alunos])
    sem = contar(res.sem_vaga)
    polo_de = {a["inst"]: a["polo"] for a in alunos}
    linhas = [{"Instituição": i, "Polo reconhecido": polo_de[i], "Lidos": n, "Roteiros 21:30": em_2130[i],
               "Roteiros 22:00": em_22[i], "Sem vaga": sem[i]} for i, n in sorted(lidos.items())]
    return pd.DataFrame(linhas)


FILL_AMARELO = PatternFill("solid", fgColor="FFFF00")
FILL_VERDE = PatternFill("solid", fgColor="2E7D32")
FONTE_TITULO = Font(bold=True, size=18)
BORDA = Border(left=Side(style="thin", color="B7B7B7"), right=Side(style="thin", color="B7B7B7"),
               top=Side(style="thin", color="B7B7B7"), bottom=Side(style="thin", color="B7B7B7"))


def escrever_roteiro_excel(ws, linha, rot, mapa_cores):
    qtd_ida, qtd_volta = contar_ida_volta(rot.alunos)

    cel = ws.cell(row=linha, column=1, value=f"{rot.id.upper()}: MOTORISTA {texto_motorista(rot, '_______________')}")
    cel.font = FONTE_TITULO
    cel.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=4)
    linha += 1

    for inst in instituicoes_do_roteiro(rot):
        fill_inst = PatternFill("solid", fgColor=mapa_cores.get(inst, "D9D9D9"))
        cel_b = ws.cell(row=linha, column=1, value=f"{inst} {rot.horario}H")
        cel_b.font = FONTE_TITULO
        cel_b.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 5):
            ws.cell(row=linha, column=c).fill = fill_inst
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=4)
        linha += 1

        for aluno in alunos_da_instituicao(rot, inst):
            eh_monitor = aluno["uid"] == rot.monitor_uid
            valores = [aluno["hora"], aluno["nome"].upper(), aluno["uso"], bairro_exibicao(aluno)]
            for c, val in enumerate(valores, start=1):
                cel_c = ws.cell(row=linha, column=c, value=val)
                cel_c.border = BORDA
                if eh_monitor:
                    cel_c.fill = FILL_VERDE
                    cel_c.font = Font(color="FFFFFF", bold=True)
                elif aluno["remanejado"]:
                    cel_c.font = Font(color="2563EB", bold=True)
                elif aluno["grupo"] in SO_UM_SENTIDO:
                    cel_c.fill = FILL_AMARELO
            linha += 1
        linha += 1

    cel_total = ws.cell(row=linha, column=1, value=f"TOTAL: IDA {qtd_ida}, RETORNO {qtd_volta}")
    cel_total.font = FONTE_TITULO
    cel_total.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=4)
    return linha + 3


def escrever_lista_excel(wb, titulo, df, larguras):
    ws = wb.create_sheet(titulo)
    for c, nome_col in enumerate(df.columns, start=1):
        ws.cell(row=1, column=c, value=nome_col).font = Font(bold=True)
    for r, linha in enumerate(df.itertuples(index=False, name=None), start=2):
        for c, val in enumerate(linha, start=1):
            ws.cell(row=r, column=c, value=val)
    for letra, w in zip("ABCDEF", larguras):
        ws.column_dimensions[letra].width = w


def gerar_excel(res, mapa_cores):
    wb = Workbook()
    ws = wb.active
    ws.title = "ROTEIRO DE IDA"

    linha_atual = 1
    for rot in res.roteiros:
        linha_atual = escrever_roteiro_excel(ws, linha_atual, rot, mapa_cores)
    for letra, w in zip("ABCD", [20, 32, 20, 24]):
        ws.column_dimensions[letra].width = w

    if res.espera_2130:
        escrever_lista_excel(wb, "LISTA DE ESPERA 21h30", tabela_alunos(res.espera_2130), [20, 32, 24, 20, 24])
    if res.sem_vaga:
        escrever_lista_excel(wb, "SEM VAGA", tabela_alunos(res.sem_vaga, com_horario=True), [20, 32, 24, 20, 24, 18])

    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()


def prioridade_veiculo(tipo):
    return {"Ônibus": 1, "Micro-ônibus": 2}.get(tipo, 3)


def configurar_frota():
    st.header("CONFIGURAÇÃO DA FROTA DE HOJE")
    st.markdown("Configure cada veículo individualmente abaixo. Você pode **adicionar novas linhas**, "
                "**excluir veículos**, alterar o número de **vagas** e preencher o **nome do motorista**.")

    if "frota_padrao" not in st.session_state:
        st.session_state.frota_padrao = pd.DataFrame([
            {"Veículo": "Ônibus", "Vagas": 46, "Horário": "21:30", "Motorista": "Cláudio"},
            {"Veículo": "Micro-ônibus", "Vagas": 32, "Horário": "21:30", "Motorista": "Van"},
            {"Veículo": "Micro-ônibus", "Vagas": 32, "Horário": "21:30", "Motorista": "Whiray"},
            {"Veículo": "Micro-ônibus", "Vagas": 30, "Horário": "21:30", "Motorista": "Romário"},
            {"Veículo": "Micro-ônibus", "Vagas": 28, "Horário": "21:30", "Motorista": "George"},
            {"Veículo": "Van", "Vagas": 15, "Horário": "22:00", "Motorista": "Anselmo"},
            {"Veículo": "Van", "Vagas": 15, "Horário": "22:00", "Motorista": "Daniel"},
            {"Veículo": "Van", "Vagas": 15, "Horário": "22:00", "Motorista": "Higor"},
        ])

    frota_df = st.data_editor(
        st.session_state.frota_padrao,
        column_config={
            "Veículo": st.column_config.SelectboxColumn("Tipo de Veículo", options=["Ônibus", "Micro-ônibus", "Van"], required=True),
            "Vagas": st.column_config.NumberColumn("Vagas Disponíveis", min_value=1, required=True),
            "Horário": st.column_config.SelectboxColumn("Horário de Retorno", options=["21:30", "22:00"], required=True),
            "Motorista": st.column_config.TextColumn("Nome do Motorista (Opcional)"),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
    )

    frota = []
    for v in frota_df.to_dict("records"):
        if pd.isna(v["Veículo"]) or pd.isna(v["Horário"]):
            continue
        mot = str(v["Motorista"]).strip() if pd.notna(v["Motorista"]) and str(v["Motorista"]).upper() != "NAN" else ""
        frota.append({"tipo": v["Veículo"], "cap": int(v["Vagas"]) if pd.notna(v["Vagas"]) else 0,
                      "horario": v["Horário"], "motorista": mot})

    por_prioridade = lambda x: prioridade_veiculo(x["tipo"])
    return (sorted([v for v in frota if v["horario"] == "21:30"], key=por_prioridade),
            sorted([v for v in frota if v["horario"] == "22:00"], key=por_prioridade))


def ler_planilha(arquivo):
    if arquivo.name.lower().endswith(".csv"):
        bruto = arquivo.getvalue()
        for enc in ("utf-8-sig", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(bruto), sep=None, engine="python", encoding=enc)
            except UnicodeDecodeError:
                continue
    return pd.read_excel(arquivo)


def encontrar_indice(colunas, palavras_chave, indice_padrao):
    for i, c in enumerate(colunas):
        chave = normalizar_chave(c)
        if any(p in chave for p in palavras_chave):
            return i
    return indice_padrao if indice_padrao < len(colunas) else 0


def mapear_colunas(colunas):
    st.sidebar.header("AJUSTE DE COLUNAS")
    with st.sidebar.expander("Mapeamento", expanded=False):
        col = {
            "hora": st.selectbox("Data/Hora", colunas, index=encontrar_indice(colunas, ["CARIMBO", "DATA HORA", "TIMESTAMP"], 0)),
            "nome": st.selectbox("Nome", colunas, index=encontrar_indice(colunas, ["NOME"], 3)),
            "inst": st.selectbox("Instituição", colunas, index=encontrar_indice(colunas, ["INSTITUICAO"], 5)),
            "uso": st.selectbox("Uso (Ida/Volta)", colunas, index=6 if len(colunas) > 6 else 0),
            "retorno": st.selectbox("Horário", colunas, index=encontrar_indice(colunas, ["RETORNO", "HORARIO"], 7)),
            "bairro": st.selectbox("Bairro/Desembarque", colunas, index=encontrar_indice(colunas, ["BAIRRO", "DESEMBARQUE"], 8)),
        }
    return col


def exibir_resumo(res, frota_2130, frota_22, alunos, descartados):
    st.subheader("RESUMO DA OPERAÇÃO")
    vagas_2130 = sum(v["cap"] for v in frota_2130)
    vagas_22 = sum(v["cap"] for v in frota_22)

    c1, c2, c3 = st.columns(3)
    c1.metric("Vagas 21:30h", vagas_2130)
    c2.metric("Demanda 21:30h", len(res.fila_2130))
    c3.metric("Saldo 21:30h", f"{vagas_2130 - len(res.fila_2130)} vagas")

    c4, c5, c6 = st.columns(3)
    c4.metric("Vagas 22:00h", vagas_22)
    c5.metric("Demanda 22:00h", len(res.fila_22))
    c6.metric("Saldo 22:00h", f"{vagas_22 - len(res.fila_22)} vagas")

    if res.sem_vaga:
        st.error(f"⚠️ ATENÇÃO: Faltou espaço para {len(res.sem_vaga)} aluno(s). "
                 "Adicione um veículo na tabela acima. Eles estão listados abaixo e na aba SEM VAGA do Excel.")
        st.subheader("Sem vaga - fora de todos os roteiros")
        st.caption("Ficaram de fora por falta de assentos, nesta ordem de chegada do formulário (os últimos a preencher).")
        st.dataframe(tabela_alunos(res.sem_vaga, com_horario=True), use_container_width=True)
    else:
        st.success("✅ Todos os estudantes foram alocados com sucesso!")

    if res.espera_2130:
        st.subheader("Lista de Espera - Vagas 21:30h")
        st.caption("Estes estudantes ficaram fora das vagas das 21:30h e foram remanejados para os carros das 22:00h, "
                   "na ordem do preenchimento do formulário.")
        st.dataframe(tabela_alunos(res.espera_2130), use_container_width=True)

    alocados = sum(len(r.alunos) for r in res.roteiros)
    nao_reconhecidas = sorted({a["inst"] for a in alunos if a["polo"] == "OUTROS"})
    if nao_reconhecidas:
        st.warning("Instituições sem polo cadastrado (tratadas como OUTROS e alocadas nas vagas livres): "
                   + ", ".join(nao_reconhecidas)
                   + ". Se alguma pertence a um polo existente, inclua-a em REGRAS_POLO no início do arquivo.")
    if alocados + len(res.sem_vaga) != len(alunos):
        st.error(f"Inconsistência: {len(alunos)} registros lidos, mas {alocados} alocados + {len(res.sem_vaga)} sem vaga.")
    with st.expander("Conferência por instituição"):
        st.caption(f"{len(alunos)} registros válidos = {alocados} em roteiros + {len(res.sem_vaga)} sem vaga. "
                   f"{len(descartados)} linha(s) descartada(s) por falta de dados.")
        st.dataframe(tabela_conferencia(alunos, res), use_container_width=True, hide_index=True)
    if descartados:
        with st.expander(f"Linhas descartadas ({len(descartados)})"):
            st.dataframe(pd.DataFrame(descartados), use_container_width=True, hide_index=True)


def estilo_linha(row):
    n = len(row)
    so_um_sentido = row["__grupo__"] in SO_UM_SENTIDO
    if row["__monitor__"]:
        return ["background-color: #2E7D32; color: #FFFFFF; font-weight: bold"] * n
    if row["__remanejado__"]:
        return [("background-color: #FFFF00; " if so_um_sentido else "") + "color: #2563EB; font-weight: bold"] * n
    if so_um_sentido:
        return ["background-color: #FFFF00; color: #000"] * n
    return [""] * n


def exibir_roteiros(res, mapa_cores):
    st.subheader("ROTEIROS PRONTOS")
    col1, col2 = st.columns(2)

    for idx, rot in enumerate(res.roteiros):
        t_ida, t_volta = contar_ida_volta(rot.alunos)
        with (col1 if idx % 2 == 0 else col2):
            with st.container(border=True):
                st.markdown(f"**{rot.id.upper()}: MOTORISTA {texto_motorista(rot, '_' * 27)}**")
                st.markdown(f"**Vagas Ocupadas: {rot.ocupacao}/{rot.cap} (Pessoas: {len(rot.alunos)}) | {t_ida} IDA, {t_volta} VOLTA**")

                for inst in instituicoes_do_roteiro(rot):
                    alunos_ord = alunos_da_instituicao(rot, inst)
                    st.markdown(
                        f"<div style='background-color:#{mapa_cores.get(inst, 'D9D9D9')};color:black;font-weight:bold;"
                        f"padding:6px;text-align:center;border-radius:4px;margin:8px 0 2px 0;'>"
                        f"{inst} {rot.horario}H</div>",
                        unsafe_allow_html=True,
                    )
                    df_exib = pd.DataFrame({
                        "Data/Hora": [a["hora"] for a in alunos_ord],
                        "Nome do Associado": [a["nome"] for a in alunos_ord],
                        "Uso do Transporte": [a["uso"] for a in alunos_ord],
                        "Bairro": [bairro_exibicao(a) for a in alunos_ord],
                        "__grupo__": [a["grupo"] for a in alunos_ord],
                        "__remanejado__": [a["remanejado"] for a in alunos_ord],
                        "__monitor__": [a["uid"] == rot.monitor_uid for a in alunos_ord],
                    })
                    df_exib.index = range(1, len(df_exib) + 1)
                    styler = df_exib.style.apply(estilo_linha, axis=1).hide(
                        subset=["__grupo__", "__remanejado__", "__monitor__"], axis="columns")
                    st.dataframe(styler, use_container_width=True)


def main():
    st.set_page_config(page_title="AJENS - ROTEIROS DE TRANSPORTE", layout="wide", page_icon="🚌")
    st.title("GERAR ROTEIROS - AJENS")

    frota_2130, frota_22 = configurar_frota()
    st.divider()

    st.header("ANALISAR PLANILHA")
    st.caption("Envie o arquivo exportado das respostas do formulário Google (.xlsx ou .csv).")
    arquivo = st.file_uploader("Envie o ficheiro Excel ou CSV de respostas", type=["xlsx", "csv"])
    if not arquivo:
        return

    df = ler_planilha(arquivo)
    colunas = list(df.columns)
    col = mapear_colunas(colunas)
    if len(set(col.values())) < len(col):
        st.error("Cada campo do mapeamento precisa apontar para uma coluna diferente. Ajuste em AJUSTE DE COLUNAS (barra lateral).")
        st.stop()

    with st.sidebar:
        if st.button("🎲 Sortear monitores novamente"):
            st.session_state.rodada_monitor = st.session_state.get("rodada_monitor", 0) + 1
    semente = f"{hashlib.md5(arquivo.getvalue()).hexdigest()}-{st.session_state.get('rodada_monitor', 0)}"

    alunos, descartados = preparar_alunos(df, col)
    res = gerar_roteiros(alunos, frota_2130, frota_22)
    sortear_monitores(res.roteiros, semente)
    mapa_cores = construir_mapa_cores(instituicoes_do_roteiro_todos(res))

    exibir_resumo(res, frota_2130, frota_22, alunos, descartados)
    exibir_roteiros(res, mapa_cores)

    st.divider()
    st.subheader("BAIXAR ROTEIROS EM PLANILHA EXCEL")
    if res.roteiros or res.sem_vaga:
        hoje = datetime.now()
        st.download_button(
            label="BAIXAR PLANILHA DE ROTEIROS (.xlsx)",
            data=gerar_excel(res, mapa_cores),
            file_name=f"{hoje.day:02d}-{hoje.month:02d}-ROTEIRO-{hoje.year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def instituicoes_do_roteiro_todos(res):
    return list(dict.fromkeys(inst for rot in res.roteiros for inst in instituicoes_do_roteiro(rot)))


if __name__ == "__main__":
    main()