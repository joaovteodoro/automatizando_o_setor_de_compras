import pandas as pd

# ABRINDO E MINERANDO AS PLANILHAS DE SOLICITAÇÕES (SDE) DO SISTEMA
def abrindo_planilhas_do_manager(caminho_planilha):
    planilha_sde = pd.read_excel(caminho_planilha)
    lista_de_sdes_sistema = planilha_sde["Sol. Web"].tolist()  # SDE VEM COM 'Compra-' na frente
    lista_de_sdes_sistema = [item.split("-")[1] for item in lista_de_sdes_sistema]  # Retira o 'Compras-'
    quantidade_sde = len(lista_de_sdes_sistema)
    print()
    return(quantidade_sde, lista_de_sdes_sistema)

# ABRINDO E MINERANDO A PLANILHA PRINCIPAL
def abrindo_planilha_principal(caminho_planilha_principal):

    planilha_principal = pd.read_excel(caminho_planilha_principal, sheet_name=1, header=2)
    return planilha_principal

# SEPARANDO A LISTA DE SOLICITAÇÕES VERMELHA(S) ('PENDENTE' E 'A FAZER') E AMARELA(S) (PARCIAL) DA PLANILHA PRINCIPAL
def salvando_sdes_da_planilha_principal(cor_sdes):

    if cor_sdes == "VERMELHAS":
        contagem = (planilha_principal["STATUS"] == "PENDENTE").sum() + (planilha_principal["STATUS"] == "A FAZER").sum()
        filtro_sdes = planilha_principal[(planilha_principal["STATUS"] == "PENDENTE") | (planilha_principal["STATUS"] == "A FAZER")]
    elif cor_sdes == "AMARELAS":
        contagem = (planilha_principal["STATUS"] == "PARCIAL").sum()
        filtro_sdes = planilha_principal[planilha_principal["STATUS"] == "PARCIAL"]

    lista_de_sdes_planilha_principal = filtro_sdes["SDE WEB"].tolist() #TRANSFORMA A TUPLA EM LISTA
    lista_de_sdes_planilha_principal = [str(elemento) for elemento in lista_de_sdes_planilha_principal]  # TRANSFORMA TODOS OS ELEMENTOS EM STR
    return(contagem, lista_de_sdes_planilha_principal)

# COMPARANDO SDES DO SISTEMA COM AS SDES DA PLANILHA 
def sde_so_esta_no_manager(lista_manager, lista_principal):
    sde_so_no_manager = lista_manager - lista_principal
    return sde_so_no_manager

# INDICANDO AS SOLICITAÇÕES LANÇADAS ERRADAS (VERMELHA QUE ESTÁ LANÇADA COMO AMARELA, E VICE E VERSA)
def sdes_de_cores_divergentes(sde_lancada_apenas_no_manager, lista_principal):
    sde_de_cor_divergente_na_principal = set()

    for sde in sde_lancada_apenas_no_manager: #confere se a sde está realmente apenas no manager, ou se foi lançada na principal com o status errado
        if sde in lista_principal:
            sde_de_cor_divergente_na_principal.add(sde) 
    
    sde_lancada_apenas_no_manager = sde_lancada_apenas_no_manager - sde_de_cor_divergente_na_principal

    return (sde_de_cor_divergente_na_principal, sde_lancada_apenas_no_manager)

# INDICA AS SOLICITAÇÕES LANÇADAS ERRADAS (DEVE(M) SER LANÇADA(S) COMO VERDE(S) ('CONCLUIDA')) 
def sde_que_devem_ser_concluidas():
    total_em_aberto_planilha = set(lista_de_sdes_vermelhas_principal + lista_de_sdes_amarelas_principal)
    total_em_aberto_sistema = set(lista_de_sdes_vermelhas_sistema + lista_de_sdes_amarelas_sistema)
    devem_ser_concluidas = total_em_aberto_planilha - total_em_aberto_sistema
    return devem_ser_concluidas

##################################################################

# ABRINDO E MINERANDO AS PLANILHAS DE SOLICITAÇÕES (SDE) DO SISTEMA
caminho_sdes_vermelhas = r"C:\\Users\\joaov\\Desktop\\PROGRAMA\\planilha_sdes_vermelhas.xlsx"
caminho_sdes_amarelas = r"C:\\Users\\joaov\\Desktop\\PROGRAMA\\planilha_sdes_amarelas.xlsx"
quantidade_sde_vermelha_sistema, lista_de_sdes_vermelhas_sistema = abrindo_planilhas_do_manager(caminho_sdes_vermelhas)
quantidade_sde_amarela_sistema, lista_de_sdes_amarelas_sistema = abrindo_planilhas_do_manager(caminho_sdes_amarelas )

# ABRINDO E MINERANDO PLANILHA PRINCIPAL
caminho_planilha_principal = r"C:\\Users\\joaov\\Desktop\\PROGRAMA\\planilha_principal.xlsm"
planilha_principal = abrindo_planilha_principal(caminho_planilha_principal)

# SEPARANDO A LISTA DE SOLICITAÇÕES VERMELHA(S) ('PENDENTE' E 'A FAZER') E AMARELA(S) (PARCIAL) DA PLANILHA PRINCIPAL
quantidade_sde_vermelha_principal, lista_de_sdes_vermelhas_principal = salvando_sdes_da_planilha_principal("VERMELHAS")    
quantidade_sde_amarela_principal, lista_de_sdes_amarelas_principal = salvando_sdes_da_planilha_principal("AMARELAS") 

# COMPARANDO SDES DO SISTEMA COM AS SDES DA PLANILHA
vermelha_so_no_manager = sde_so_esta_no_manager(set(lista_de_sdes_vermelhas_sistema), set(lista_de_sdes_vermelhas_principal))
amarela_so_no_manager = sde_so_esta_no_manager(set(lista_de_sdes_amarelas_sistema), set(lista_de_sdes_amarelas_principal))

# INDICANDO AS SOLICITAÇÕES LANÇADAS ERRADAS (VERMELHA QUE ESTÁ LANÇADA COMO AMARELA, E VICE E VERSA)
amarela_que_deveria_ser_vermelha, vermelha_so_no_manager = sdes_de_cores_divergentes(vermelha_so_no_manager, lista_de_sdes_amarelas_principal)
vermelha_que_deveria_ser_amarela, amarela_so_no_manager = sdes_de_cores_divergentes(amarela_so_no_manager, lista_de_sdes_vermelhas_principal)

# INDICA AS SOLICITAÇÕES LANÇADAS ERRADA (DEVEM SER LANÇADA COMO VERDE (CONCLUIDA)) 
devem_ser_concluidas = sde_que_devem_ser_concluidas()

# APRESENTA AS INFORMAÇÕES FINAIS
def informacoes_gerais():
    info = f"""
    <h2>SISTEMA:</h2>
    <ul>
        <li>SDE(S) <strong><span style="color: red;">VERMELHA</span></strong>: {quantidade_sde_vermelha_sistema}</li>
        <li>{lista_de_sdes_vermelhas_sistema}</li>
    </ul>
    
    <ul>
        <li>SDE(S) <strong><span style="color: #B8860B;">AMARELA</span></strong>: {quantidade_sde_amarela_sistema}</li>
        <li>{lista_de_sdes_amarelas_sistema}</li>
    </ul>

    <h2>PLANILHA:</h2>
    <ul>
        <li>SDE(S) <strong><span style="color: red;">VERMELHA</span></strong>: {quantidade_sde_vermelha_principal}</li>
        <li>{lista_de_sdes_vermelhas_principal}</li>
    </ul>
    
    <ul>
        <li>SDE(S) <strong><span style="color: #B8860B;">AMARELA</span></strong>: {quantidade_sde_amarela_principal}</li>
        <li>{lista_de_sdes_amarelas_principal}</li>
    </ul>

    <h2>INFORMAÇÕES GERAIS:</h2>
    <ul>
        <li>A(s) SDE(s) <strong>{amarela_que_deveria_ser_vermelha}</strong> deve(m) ser alterada(s) para <strong><span style="color: red;">VERMELHO</span></strong></li>
        <li>A(s) SDE(s) <strong>{vermelha_que_deveria_ser_amarela}</strong> deve(m) ser alterada(s) para <strong><span style="color: #B8860B;">AMARELO</span></strong></li>
        <li>A(s) SDE(s) <strong>{devem_ser_concluidas}</strong> deve(m) ser alterada(s) para <strong><span style="color: green;">VERDE</span></strong></li>
    </ul>

    <ul>
        <li>A(s) SDE(s) <strong><span style="color: red;">VERMELHA(S)</span></strong> que deve(m) ser incluida(s) na planilha são: <strong>{vermelha_so_no_manager}</strong></li>
        <li>A(s) SDE(s) <strong><span style="color: #B8860B;">AMARELA(S)</span></strong> que deve(m) ser incluida(s) na planilha são: <strong>{amarela_so_no_manager}</strong></li>
    </ul>
    """
    return info