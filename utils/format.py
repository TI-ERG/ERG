def formatar_valor(valor, moeda=False):
    formatado = "{:,.2f}".format(valor)
    if moeda:
        return "R$ " + formatado.replace(",", "v").replace(".", ",").replace("v", ".")
    else:
        return formatado.replace(",", "v").replace(".", ",").replace("v", ".").replace(",00", "")
    
def arredondar_decimais(df, cols):
    df[cols] = (df[cols].astype(float).round(2))
    return df    

def copiar_estilo(ws, origem, destino):
    for col in range(1, ws.max_column + 1):
        cel_origem = ws.cell(row=origem, column=col)
        cel_destino = ws.cell(row=destino, column=col)
        cel_destino._style = cel_origem._style
        