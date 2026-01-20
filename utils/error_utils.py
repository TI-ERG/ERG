class ErroDePeriodo(Exception): 
    def __init__(self, data, mensagem="O período do arquivo não confere com o período determinado"): 
        super().__init__(f"{mensagem}: {data.strftime("%m/%Y")}")


class LayoutInesperado(Exception):
    pass
