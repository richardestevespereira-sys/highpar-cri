# -*- coding: utf-8 -*-
"""
ATUALIZADO1 — diagnostico.py: rastreador AUTÔNOMO do dossiê documental por emissão.

Para cada CRI monitorado, existe um conjunto-padrão de documentos que a mesa
PRECISA ter (é deles que saem covenants, garantias/colaterais e o lastro).
Este script varre as pastas do monitoramento SOZINHO (por convenção de nomes),
marca cada item do dossiê como presente/ausente, mede a recência e grava o
resultado em cadastro/<id>.json (bloco "dossie"). O atualizar_dados.py leva o
dossiê para o portal, que exibe o diagnóstico e os avisos.

Uso:  python diagnostico.py            (roda para todas as emissões do cadastro)
Rotina: chamado automaticamente pelo atualizar_mes.bat antes do atualizar_dados.
"""
import os, re, json, glob, unicodedata, datetime as dt

RAIZ = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
IGNORAR_DIRS = {'.git', '.agents', '__pycache__', 'node_modules', '.venv',
                'Arquivos Antigos', 'Sintese Macro', 'site'}

def sem_acento(s):
    return unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().lower()

# ---------------------------------------------------------------------------
# DOSSIÊ-PADRÃO: o que TODO CRI deveria ter, por quê, e como reconhecer o
# arquivo pelo nome. "obrigatorio": False = condicional (depende da operação).
# ---------------------------------------------------------------------------
DOSSIE_PADRAO = [
 {'id':'termo_securitizacao','nome':'Termo de Securitização (TS)','cat':'Jurídico','obrigatorio':True,
  'motivo':'Documento-mãe: define séries, taxas, amortização, covenants, garantias e eventos de vencimento antecipado.',
  'period':None,'re':r'termo.*securitiza|(^|[_\s-])ts([_\s.\-]|$)|adit.*ts'},
 {'id':'contrato_cessao','nome':'Contrato de Cessão / Cessão Fiduciária','cat':'Garantias','obrigatorio':True,
  'motivo':'Formaliza a cessão dos recebíveis (o lastro). Sem ele não se comprova a titularidade da carteira.',
  'period':None,'re':r'cess[aã]o'},
 {'id':'alienacao_fiduciaria','nome':'Alienação Fiduciária (imóveis/quotas) registrada','cat':'Garantias','obrigatorio':True,
  'motivo':'Colateral real da operação (AFI/AFP). A versão REGISTRADA (RGI/RTD) é a que vale contra terceiros.',
  'period':None,'re':r'aliena[cç][aã]o|af[_\s]*quotas|afi|afp|fiduci'},
 {'id':'cci','nome':'CCI / Escritura de emissão','cat':'Jurídico','obrigatorio':False,
  'motivo':'Cédula que representa os créditos imobiliários cedidos.',
  'period':None,'re':r'(^|[_\s-])cci([_\s.\-]|$)'},
 {'id':'matricula','nome':'Matrícula(s) do(s) imóvel(is)','cat':'Garantias','obrigatorio':True,
  'motivo':'Comprova o registro do gravame (AF) na matrícula-mãe/filhas. Base do LTV.',
  'period':12,'re':r'matr[ií]cula'},
 {'id':'cnd','nome':'CNDs da Cedente/SPE (federal, trabalhista, municipal)','cat':'Covenants','obrigatorio':True,
  'motivo':'Covenant típico: CND vencida/positiva pode dar direito a pedir ANTECIPAÇÃO do recebimento da dívida.',
  'period':6,'re':r'(^|[_\s-])cnd|certid[aã]o.*(negativa|d[eé]bito)'},
 {'id':'apolice_seguro','nome':'Apólice de seguro + endosso (beneficiária: Securitizadora)','cat':'Covenants','obrigatorio':True,
  'motivo':'Sem endosso vigente, um sinistro na obra/imóvel não protege o patrimônio separado.',
  'period':12,'re':r'ap[oó]lice|endosso|seguro'},
 {'id':'declaracoes','nome':'Declarações periódicas (adimplência, recompra/multa)','cat':'Covenants','obrigatorio':True,
  'motivo':'Obrigações declaratórias do TS; a ausência gera apontamento do Agente Fiduciário.',
  'period':12,'re':r'declara[cç][aã]o'},
 {'id':'relatorio_servicer','nome':'Relatório mensal do servicer (carteira/fechamento)','cat':'Monitoramento','obrigatorio':True,
  'motivo':'Fonte primária da carteira: recebidas, a receber, inadimplência, aging, vendas/estoque.',
  'period':1,'re':r'fechamento|consolidado_(mensal|receb)|monitoramento(?!_obra)|carteira'},
 {'id':'medicao_obra','nome':'Medição/monitoramento mensal de obra','cat':'Monitoramento','obrigatorio':False,
  'motivo':'Se há obra, a medição libera recursos e alimenta o cronograma físico (%). Sem ela, obra às cegas.',
  'period':1,'re':r'medi[cç][aã]o|monitoramento_obra|obra_|_obra'},
 {'id':'extrato_conta','nome':'Extratos da conta centralizadora','cat':'Monitoramento','obrigatorio':True,
  'motivo':'Prova do caixa: recebimentos da carteira, despesas, pagamentos ao CRI e fundos (reserva/obra).',
  'period':1,'re':r'extrato|consultapagamentos|mov[\s_]*fin'},
 {'id':'balancete_spe','nome':'Balancete da SPE/devedora','cat':'Monitoramento','obrigatorio':False,
  'motivo':'Saúde financeira do devedor; antecipa estresse antes de virar inadimplência.',
  'period':3,'re':r'balancete'},
 {'id':'agt','nome':'Assembleias (AGT) e instruções de voto','cat':'Jurídico','obrigatorio':False,
  'motivo':'Alterações de condições, waivers e renegociações passam por AGT — mudam covenants e garantias.',
  'period':None,'re':r'assembleia|(^|[_\s-])agt|instru[cç][aã]o de voto'},
 {'id':'subscricao','nome':'Boletins de subscrição / integralização','cat':'Posição','obrigatorio':False,
  'motivo':'Comprova a posição detida (qtd por série) — base do % do ativo e do valor da posição Highpar.',
  'period':None,'re':r'subscri|integraliza|aporte'},
]

MES_RE = re.compile(r'(20\d{2})[./\\-]?(0[1-9]|1[0-2])|(?:^|[^0-9])(0[1-9]|1[0-2])\.(2\d)(?:[^0-9]|$)')

def recencia(caminho):
    """Tenta extrair AAAA-MM do caminho (padrões 2026/04.26, 04.26, 2026-04...)."""
    melhores = []
    for m in MES_RE.finditer(sem_acento(caminho)):
        if m.group(1): melhores.append(f"{m.group(1)}-{m.group(2)}")
        elif m.group(3): melhores.append(f"20{m.group(4)}-{m.group(3)}")
    return max(melhores) if melhores else None

def varrer_arquivos(pastas):
    out = []
    for p in pastas:
        base = os.path.join(RAIZ, p) if not os.path.isabs(p) else p
        if not os.path.isdir(base): continue
        for r, dirs, files in os.walk(base):
            dirs[:] = [d for d in dirs if d not in IGNORAR_DIRS]
            for f in files:
                if f.startswith('~$'): continue
                out.append(os.path.relpath(os.path.join(r, f), RAIZ))
    return out

def diagnosticar(cad_path):
    with open(cad_path, encoding='utf-8') as fh:
        cad = json.load(fh)
    dcfg = cad.get('dossie_config') or {}
    pastas = dcfg.get('pastas_busca') or []
    arquivos = varrer_arquivos(pastas)
    nomes = [(a, sem_acento(os.path.basename(a))) for a in arquivos]
    itens, presentes_obr, total_obr = [], 0, 0
    for item in DOSSIE_PADRAO:
        if item['id'] in (dcfg.get('nao_aplica') or []):
            continue
        rx = re.compile(item['re'])
        hits = [a for a, n in nomes if rx.search(n)]
        rec = max((recencia(h) or '' for h in hits), default='') or None
        st = 'presente' if hits else ('na_nuvem' if item['id'] in (dcfg.get('na_nuvem') or []) else 'ausente')
        # presença com periodicidade: se o mais recente está velho, degrada
        atraso = None
        if st == 'presente' and item['period'] and rec:
            a, m = map(int, rec.split('-')); hoje = dt.date.today()
            atraso = (hoje.year - a) * 12 + (hoje.month - m)
            if atraso > item['period'] + 1: st = 'desatualizado'
        if item['obrigatorio']:
            total_obr += 1
            if st == 'presente': presentes_obr += 1
        itens.append({'id': item['id'], 'nome': item['nome'], 'cat': item['cat'],
                      'motivo': item['motivo'], 'obrigatorio': item['obrigatorio'],
                      'period': item['period'], 'status': st, 'qtd': len(hits),
                      'mais_recente': rec, 'exemplo': (sorted(hits)[-1] if hits else None)})
    score = round(100 * presentes_obr / total_obr) if total_obr else None
    cad['dossie'] = {'gerado': dt.datetime.now().isoformat(timespec='minutes'),
                     'pastas_busca': pastas, 'score_documental': score,
                     'obrigatorios_presentes': f'{presentes_obr}/{total_obr}', 'itens': itens}
    with open(cad_path, 'w', encoding='utf-8') as fh:
        json.dump(cad, fh, ensure_ascii=False, indent=2)
    return cad['id'], score, itens

def main():
    achados = glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cadastro', '*.json'))
    if not achados:
        raise SystemExit('Nenhum cadastro/<id>.json encontrado.')
    for c in sorted(achados):
        eid, score, itens = diagnosticar(c)
        falt = [i['nome'] for i in itens if i['obrigatorio'] and i['status'] != 'presente']
        print(f'{eid}: score documental {score if score is not None else "-"}/100')
        for i in itens:
            print(f"   [{i['status']:>13}] {i['nome']}" + (f" (últ. {i['mais_recente']}, {i['qtd']} arq.)" if i['qtd'] else ''))
        if falt: print('   FALTANDO (obrigatórios):', '; '.join(falt))
    print('Dossiês gravados no cadastro. Rode: python atualizar_dados.py')

if __name__ == '__main__':
    main()
