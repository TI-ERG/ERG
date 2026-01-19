import datetime
import math
import calendar

#--------------------------------------------------------
# Retorna a quantidade de semanas no mês.
# input: string da data
# output: número da semana
#--------------------------------------------------------
def semanas_no_mes(data):
    # garante que é datetime.date
    if isinstance(data, str):
        data = datetime.datetime.strptime(data, "%Y-%m-%d").date()

    ano = data.year
    mes = data.month

    # primeiro e último dia do mês
    primeiro = datetime.date(ano, mes, 1)
    ultimo = datetime.date(ano, mes, calendar.monthrange(ano, mes)[1])

    # weekday(): segunda=0 ... domingo=6
    inicio_semana = primeiro.weekday()  # deslocamento da primeira segunda
    dias_no_mes = (ultimo - primeiro).days + 1

    # total de dias considerando o deslocamento até a primeira segunda
    total_dias = inicio_semana + dias_no_mes

    # semanas completas + incompletas
    return math.ceil(total_dias / 7)


#--------------------------------------------------------
# Retorna a semana no mês que a data pertence
# input: string da data
# output: número da semana
#--------------------------------------------------------
def semana_do_mes(data):
    return (data.day - 1) // 7 + 1

#--------------------------------------------------------
# Retorna a semana por extenso pelo número.
# input: número da semana 
# output: string extenso
#--------------------------------------------------------
def semana_extenso_numero(numero):
    nomes = [
        "Primeira Semana",
        "Segunda Semana",
        "Terceira Semana",
        "Quarta Semana",
        "Quinta Semana",
        "Sexta Semana",
        "Sétima Semana"
    ]
    if 1 <= numero <= 7:
        return nomes[numero - 1]
    return "Semana Inválida"

#--------------------------------------------------------
# Retorna a semana por extenso pela data. 
# input: string da data 
# output: número da semana
#--------------------------------------------------------
def semana_extenso_data(data):
    num = semanas_no_mes(data)
    return semana_extenso_numero(num)

#--------------------------------------------------------
# Retorna o dia da semana por extenso pela data. 
# input: string da data 
# output: string dia da semana
#--------------------------------------------------------
def dia_da_semana(data):
    # garante que é datetime.date
    if isinstance(data, str):
        data = datetime.datetime.strptime(data, "%Y-%m-%d").date()

    nomes = [
        "Segunda-feira",
        "Terça-feira",
        "Quarta-feira",
        "Quinta-feira",
        "Sexta-feira",
        "Sábado",
        "Domingo",
    ]

    return nomes[data.weekday()]
