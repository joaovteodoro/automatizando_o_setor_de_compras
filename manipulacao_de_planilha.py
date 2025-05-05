import pandas as pd

print()
print("SISTEMA:")
print("----------------")

# SALVANDO A LISTA DE SDES VERMELHA(S) NO SISTEMA
caminho_sde_vermelhas = r"C:\\Users\\joaov\\Desktop\\PROGRAMA\\planilha_sdes_vermelhas.xlsx"
planilha_sde_vermelhas = pd.read_excel(caminho_sde_vermelhas)
lista_de_sdes_vermelhas_sistema = planilha_sde_vermelhas["Sol. Web"].tolist()  # SDE VEM COM 'Compra-' na frente
lista_de_sdes_vermelhas_sistema = [item.split("-")[1] for item in lista_de_sdes_vermelhas_sistema]  # Retira o 'Compras-'
print(f"SDE(s) VERMELHA(S): {len(lista_de_sdes_vermelhas_sistema)}")
print(lista_de_sdes_vermelhas_sistema)

# SALVANDO A LISTA DE SDES AMARELA(S) NO SISTEMA
caminho_sdes_amarelas = r"C:\\Users\\joaov\\Desktop\\PROGRAMA\\planilha_sdes_amarelas.xlsx"
planilha_sdes_amarelas = pd.read_excel(caminho_sdes_amarelas)
lista_de_sdes_amarelas_sistema = planilha_sdes_amarelas["Sol. Web"].tolist()  # SDE VEM COM 'Compra-' na frente
lista_de_sdes_amarelas_sistema = [item.split("-")[1] for item in lista_de_sdes_amarelas_sistema]  # Retira o 'Compras-'
print(f"SDE(s) AMARELA(S): {len(lista_de_sdes_amarelas_sistema)}")
print(lista_de_sdes_amarelas_sistema)

print()
print("PLANILHA:")
print("----------------")
#############################################

# ABRINDO PLANILHA PRINCIPAL
caminho_principal = r"C:\\Users\\joaov\\Desktop\\PROGRAMA\\planilha_principal.xlsm"
planilha_principal = pd.read_excel(caminho_principal, sheet_name=1, header=2)
# sheet_name = 1 - Seleciona a segunda aba da planilha
# header mostra em qual linha começa o cabeçalho

# SALVANDO A LISTA DE SDES VERMELHA(S) NA PLANILHA ('PENDENTE' E 'A FAZER')
contagem_pendente = (planilha_principal["STATUS"] == "PENDENTE").sum() + (planilha_principal["STATUS"] == "A FAZER").sum()
print(f"SDE(s) VERMELHA(S): {contagem_pendente}")
filtro_pendentes = planilha_principal[planilha_principal["STATUS"] == "PENDENTE"]
filtro_a_fazer = planilha_principal[planilha_principal["STATUS"] == "A FAZER"]
lista_de_sdes_vermelhas_principal = filtro_pendentes["SDE WEB"].tolist() + filtro_a_fazer["SDE WEB"].tolist()
lista_de_sdes_vermelhas_principal = [
    str(elemento) for elemento in lista_de_sdes_vermelhas_principal
]  # TRANSFORMA TODOS OS ELEMENTOS EM STR
print(lista_de_sdes_vermelhas_principal)

# SALVANDO A LISTA DE SDES AMARELA(S) NA PLANILHA ('PARCIAL')
contagem_parcial = (planilha_principal["STATUS"] == "PARCIAL").sum()
print(f"SDE(s) AMARELA(S): {contagem_parcial}")
filtro_parcial = planilha_principal[planilha_principal["STATUS"] == "PARCIAL"]
lista_de_sdes_amarelas_principal = filtro_parcial["SDE WEB"].tolist()
lista_de_sdes_amarelas_principal = [str(elemento) for elemento in lista_de_sdes_amarelas_principal]
print(lista_de_sdes_amarelas_principal)
print()

# COMPARANDO SDES DO SISTEMA COM AS SDES DA PLANILHA
vermelha_so_no_sistema = set(lista_de_sdes_vermelhas_sistema) - set(lista_de_sdes_vermelhas_principal)
amarela_so_no_sistema = set(lista_de_sdes_amarelas_sistema) - set(lista_de_sdes_amarelas_principal)

# RESOLVENDO QUANDO UMA SDE VERMELHA E)STÁ LANÇADA COMO AMARELA, E VICE E VERSA
amarela_que_deveria_ser_vermelha = set()
for sde in vermelha_so_no_sistema:
    if sde in lista_de_sdes_amarelas_principal:
        amarela_que_deveria_ser_vermelha.add(
            sde
        )  # outra forma de trabalhar era separando a união entre a listaDeSdeAmarelasPrincipal e VermelhoSoNoSistema
vermelha_so_no_sistema = vermelha_so_no_sistema - amarela_que_deveria_ser_vermelha

vermelha_que_deveria_ser_amarela = set()
for sde in amarela_so_no_sistema:
    if sde in lista_de_sdes_vermelhas_principal:
        vermelha_que_deveria_ser_amarela.add(sde)
amarela_so_no_sistema = amarela_so_no_sistema - vermelha_que_deveria_ser_amarela

# RESOLVENDO QUANDO UMA SDE DEVERIA ESTAR CONCLUIDA NA PLANILHA, E NÃO ESTÁ
total_em_aberto_planilha = set(lista_de_sdes_vermelhas_principal + lista_de_sdes_amarelas_principal)
total_em_aberto_sistema = set(lista_de_sdes_vermelhas_sistema + lista_de_sdes_amarelas_sistema)
devem_ser_concluidas = total_em_aberto_planilha - total_em_aberto_sistema

print("INFORMAÇÕES GERAIS:")
print("----------------")
print(f"A(s) SDE(s) {amarela_que_deveria_ser_vermelha} deve(m) ser alterada(s) para VERMELHO")
print(f"A(s) SDE(s) {vermelha_que_deveria_ser_amarela} deve(m) ser alterada(s) para AMARELO")
print(f"A(s) SDE(s) {devem_ser_concluidas} deve(m) ser alterada(s) para CONCLUIDO")
print()
print(f"A(s) SDE(s) VERMELHA(S)) que deve(m) ser incluida(s) na planilha são: {vermelha_so_no_sistema}")
print(f"A(s) SDE(s) AMARELA(S) que deve(m) ser incluida(s) na planilha são: {amarela_so_no_sistema}")
print()
