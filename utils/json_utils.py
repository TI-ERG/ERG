import json

def ler_json(caminho):
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)

def salvar_json(df, caminho):
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(df, f, indent=4, ensure_ascii=False)