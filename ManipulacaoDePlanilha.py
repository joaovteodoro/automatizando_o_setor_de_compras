import pandas as pd

print()
print('SISTEMA:')
print('----------------')

#SALVANDO A LISTA DE SDES VERMELHA(S) NO SISTEMA
CaminhoSdeVermelhas = r'C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeVermelha.xlsx'
PlanilhaSdeVermelhas = pd.read_excel(CaminhoSdeVermelhas)
ListaDeSdeVermelhasSistema = PlanilhaSdeVermelhas['Sol. Web'].tolist() #SDE VEM COM 'Compra-' na frente
ListaDeSdeVermelhasSistema = [item.split('-')[1] for item in ListaDeSdeVermelhasSistema] #Retira o 'Compras-'
print(f'SDE(s) VERMELHA(S): {len(ListaDeSdeVermelhasSistema)}')
print(ListaDeSdeVermelhasSistema)

#SALVANDO A LISTA DE SDES AMARELA(S) NO SISTEMA
CaminhoSdeAmarelas = r'C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeAmarela.xlsx'
PlanilhaSdeAmarelas = pd.read_excel(CaminhoSdeAmarelas)
ListaDeSdeAmarelasSistema = PlanilhaSdeAmarelas['Sol. Web'].tolist() #SDE VEM COM 'Compra-' na frente
ListaDeSdeAmarelasSistema = [item.split('-')[1] for item in ListaDeSdeAmarelasSistema] #Retira o 'Compras-'
print(f'SDE(s) AMARELA(S): {len(ListaDeSdeAmarelasSistema)}')
print(ListaDeSdeAmarelasSistema)

print()
print('PLANILHA:')
print('----------------')
#############################################

#ABRINDO PLANILHA PRINCIPAL
CaminhoPrincipal = r'C:\\Users\\joaov\\Desktop\\PROGRAMA\\principal.xlsm'
PlanilhaPrincipal = pd.read_excel(CaminhoPrincipal, sheet_name=1, header=2)
    # sheet_name = 1 - Seleciona a segunda aba da planilha
    # header mostra em qual linha começa o cabeçalho

#SALVANDO A LISTA DE SDES VERMELHA(S) NA PLANILHA ('PENDENTE' E 'A FAZER') 
ContagemPendente = (PlanilhaPrincipal['STATUS'] == 'PENDENTE').sum() + (PlanilhaPrincipal['STATUS'] == 'A FAZER').sum()
print(f'SDE(s) VERMELHA(S): {ContagemPendente}')
FiltroPendentes = PlanilhaPrincipal[PlanilhaPrincipal['STATUS'] == 'PENDENTE']
FiltroAFazer = PlanilhaPrincipal[PlanilhaPrincipal['STATUS'] == 'A FAZER']
ListaDeSdeVermelhasPrincipal = FiltroPendentes['SDE WEB'].tolist() + FiltroAFazer['SDE WEB'].tolist()
ListaDeSdeVermelhasPrincipal = [str(elemento) for elemento in ListaDeSdeVermelhasPrincipal] # TRANSFORMA TODOS OS ELEMENTOS EM STR
print(ListaDeSdeVermelhasPrincipal)

#SALVANDO A LISTA DE SDES AMARELA(S) NA PLANILHA ('PARCIAL')
ContagemParcial = (PlanilhaPrincipal['STATUS'] == 'PARCIAL').sum()
print(f'SDE(s) AMARELA(S): {ContagemParcial}')
FiltroParcial = PlanilhaPrincipal[PlanilhaPrincipal['STATUS'] == 'PARCIAL']
ListaDeSdeAmarelasPrincipal = FiltroParcial['SDE WEB'].tolist()
ListaDeSdeAmarelasPrincipal = [str(elemento) for elemento in ListaDeSdeAmarelasPrincipal]
print(ListaDeSdeAmarelasPrincipal)
print()

#COMPARANDO SDES DO SISTEMA COM AS SDES DA PLANILHA
VermelhaSoNoSistema = set(ListaDeSdeVermelhasSistema) - set(ListaDeSdeVermelhasPrincipal) 
AmarelaSoNoSistema = set(ListaDeSdeAmarelasSistema) - set(ListaDeSdeAmarelasPrincipal)

# RESOLVENDO QUANDO UMA SDE VERMELHA E)STÁ LANÇADA COMO AMARELA, E VICE E VERSA 
AmarelaQueDeveriaSerVermelha = set()
for sde in VermelhaSoNoSistema:
    if sde in ListaDeSdeAmarelasPrincipal:
        AmarelaQueDeveriaSerVermelha.add(sde) #outra forma de trabalhar era separando a união entre a listaDeSdeAmarelasPrincipal e VermelhoSoNoSistema
VermelhaSoNoSistema = VermelhaSoNoSistema - AmarelaQueDeveriaSerVermelha

VermelhaQueDeveriaSerAmarela = set()
for sde in AmarelaSoNoSistema:
    if sde in ListaDeSdeVermelhasPrincipal:
        VermelhaQueDeveriaSerAmarela.add(sde)
AmarelaSoNoSistema = AmarelaSoNoSistema - VermelhaQueDeveriaSerAmarela

#RESOLVENDO QUANDO UMA SDE DEVERIA ESTAR CONCLUIDA NA PLANILHA, E NÃO ESTÁ
TotalEmAbertoPlanilha = set(ListaDeSdeVermelhasPrincipal + ListaDeSdeAmarelasPrincipal)
TotalEmAbertoSistema =set(ListaDeSdeVermelhasSistema + ListaDeSdeAmarelasSistema)
DevemSerConcluidas = TotalEmAbertoPlanilha - TotalEmAbertoSistema

print('INFORMAÇÕES GERAIS:')
print('----------------')
print(f'A(s) SDE(s) {AmarelaQueDeveriaSerVermelha} deve(m) ser alterada(s) para VERMELHO')
print(f'A(s) SDE(s) {VermelhaQueDeveriaSerAmarela} deve(m) ser alterada(s) para AMARELO')
print(f'A(s) SDE(s) {DevemSerConcluidas} deve(m) ser alterada(s) para CONCLUIDO')
print()
print(f'A(s) SDE(s) VERMELHA(S)) que deve(m) ser incluida(s) na planilha são: {VermelhaSoNoSistema}')
print(f'A(s) SDE(s) AMARELA(S) que deve(m) ser incluida(s) na planilha são: {AmarelaSoNoSistema}')
print()
