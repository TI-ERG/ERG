import json

def ler_json(arquivo):
    with open(arquivo, "r", encoding="utf-8") as f:
        return json.load(f)

def salvar_json(df, arquivo):
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump(df, f, indent=4, ensure_ascii=False)