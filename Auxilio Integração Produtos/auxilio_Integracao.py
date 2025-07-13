from openpyxl import load_workbook
class Integracao():
    def __init__(self,tipo):
        self.tipo = tipo
        self._valor = []
        
    @property
    def valor(self):
        return self._valor
    
    @valor.setter
    def valor(valo,self):
        self.valor.append(valo)
    
class Planilha():
    def __init__(self,nome):
        self.planilha = nome

def valida_Planilha(nome):
    try:
        planilha = load_workbook(nome+".xlsx")
        return Planilha(planilha)
    except Exception as e:
        print(e)

def adiciona_Planilha(planilhas):
    plan = None
    while not plan:
        plan=input("Informe o nome da planilha: ")
        if plan in planilhas:
            print("Planilha ja importada")
            plan = ""
        elif not valida_Planilha(plan):
            plan = ""
        else:
            planilhas.append(plan)
            return valida_Planilha(plan)
        
planilhas = []
integracao = None
while not integracao:
    integracao = Integracao(tipo=int(input("Informe qual a integração:\n1-Ifood\n")))

if input("Adicionar planilha de produto?(S/N): ").upper() == "S":
    produto = adiciona_Planilha(planilhas)
    integracao.valor.append("Produto")
    
if input("Adicionar planilha de Pizza?(S/N): ").upper() == "S":
    pizza = adiciona_Planilha(planilhas)
    integracao.valor.append("Pizza")

if input("Adicionar planilha de Adicionais?(S/N): ").upper() == "S":
    adicional = adiciona_Planilha(planilhas)
    integracao.valor.append("Adicional")

