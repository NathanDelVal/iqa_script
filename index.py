import pandas as pd
import sys
import re
import params
from datetime import datetime
import dateparser
import argparse
from unidecode import unidecode

target_month = 8

# --- argparse setup ---
parser = argparse.ArgumentParser(description="Process IQA Excel files.")
parser.add_argument('input_files', nargs='+', help='Input Excel files')
parser.add_argument('--referencia', default=int(datetime.now().__format__('%m')), help='Coloque o mês referente')
parser.add_argument('--output', default=f"IQA_{datetime.now().__format__('%m-%Y')}.xlsx", help='Output Excel file name')
parser.add_argument('--log', default="log.txt", help='Log file name')
args = parser.parse_args()

sheet_reference = {"Plan_Conc":"05-PLN_AMT_VRF", "Dados_Conc":"08-RST_ANL_VRF"}
sheet_reference_name = [x for x in sheet_reference.values()]
conc_dict = {
 1: 'set', 2: 'out', 3: 'nov', 4: 'dez',
 5: 'jan', 6: 'fev', 7: 'mar', 8: 'abr',
 9: 'mai', 10: 'jun', 11: 'jul', 12: 'ago'
}

blocos = {"a": params.bloco_a, "b": params.bloco_b, "c": params.bloco_c}

iqa_sheets = {"Plan_Conc": [], "Dados_Conc": []}
sheets_conc = {}

if args.input_files:
    for x in args.input_files:
        if re.search(params.file_pattern, x) is None:
            raise Exception(f"O arquivo {x} não está no padrão correto! Favor padronizar. Ex: ...(a).xlsx")
        try:
            df = pd.read_excel(x, sheet_name=sheet_reference_name)
            for sheet in sheet_reference_name:
                final_col = []
                prepend = {
                    "Relatório": [f"{args.referencia}{''.join(re.findall(params.file_pattern, x)).upper()}"] * len(df[sheet]),
                    "Mês Ref": [f"{conc_dict[args.referencia]}/{datetime.now().__format__('%y')}"] * len(df[sheet]),
                    "Bloco": ["".join(re.findall(params.file_pattern, x)).upper()] * len(df[sheet]),
                    "Empresa": [blocos["".join(re.findall(params.file_pattern, x)).lower()]] * len(df[sheet])
                }
                if sheet == "08-RST_ANL_VRF":
                    sheet_cols = df[sheet].columns.tolist()
                    lab_col = "".join([x for x in sheet_cols if unidecode(x.lower().replace(" ","_")) == "laboratorio_analise"])
                    tipo_col = "".join([x for x in sheet_cols if unidecode(x.lower().replace(" ","_")) == "tipo"])
                    lab_col = df[sheet][lab_col].values.tolist() if lab_col in sheet_cols else None
                    tipo_col = df[sheet][tipo_col].values.tolist() if tipo_col in sheet_cols else None
                    if lab_col is not None:
                        prepend["Tipo Laboratório"] = ["Interno" if (re.search("RMM - LT - Bioagri", str(v), re.IGNORECASE) and not pd.isna(v)) else "Externo" if not pd.isna(v) else v for v in lab_col]
                    if tipo_col is not None:
                        prepend["Tipo Laboratório"] = ["Interno" if (re.search("externo", str(v), re.IGNORECASE) is None and not pd.isna(v)) else "Externo" if not pd.isna(v) else v for v in tipo_col]
                prepend = pd.DataFrame(prepend)        
                cols_to_drop = [col for col in df[sheet].columns if re.search("Unnamed", col)] + ["Concessão", "Empresa"]
                df[sheet] = df[sheet].drop([col for col in cols_to_drop if col in df[sheet].columns], axis=1)
                iqa_sheets["".join([k for k, v in sheet_reference.items() if v == sheet])].append(pd.concat([prepend, df[sheet]], axis=1))
            print("✅ Excel file loaded successfully!")
        except FileNotFoundError:
            print(f"❌ File not found: {x}")
            exit()
else:
    raise Exception("Por favor, inclua os nomes dos arquivos nos argumentos! Ex: python index.py [arquivo1] [arquivo2]...")

for k, v in iqa_sheets.items():
    intersection_columns_list = list(set.intersection(*[set([unidecode(col.lower().replace(" ","_")) for col in df.columns]) for df in v]))
    for sheet in v:
        result = [col for col in sheet.columns.to_list() if unidecode(col.lower().replace(" ","_")) in intersection_columns_list]
        sheet = sheet[result]
    iqa_sheets[k] = pd.concat(v)

relatorio = [f"{x}A" for x in range(1,13)] + [f"{x}B" for x in range(1,13)] + [f"{x}C" for x in range(1,13)]
empresa = [params.bloco_a] * 12 + [params.bloco_b] * 12 + [params.bloco_c] * 12
bloco = ["A"] * 12 + ["B"] * 12 + ["C"] * 12

mes = []

for x in range(1, 13):
    year = dateparser.parse("last year").__format__("%y")
    if x > 4:
        year = datetime.now().__format__("%y")
    mes.append(f"{conc_dict[x]}/{year}")

mes *= 3

iqa_sheets["Lista"] = pd.DataFrame({
    "Relatório": relatorio,
    "Empresa": empresa,
    "Bloco": bloco,
    "Mês": mes
})

with pd.ExcelWriter(args.output) as writer:
    for sheet, df in iqa_sheets.items():
        df.to_excel(writer, index=False, sheet_name=sheet)