import numpy as np
import matplotlib.pyplot as plt
import pandas as pd


def carregar_dados(nome_arquivo, planilha):
    """Carrega os dados do arquivo Excel."""
    dados = pd.DataFrame(pd.read_excel(nome_arquivo, sheet_name=planilha)).sort_values("Date")
    return dados


def calcular_estatisticas_retorno(dados):
    """Calcula a média e a volatilidade dos retornos diários."""
    retorno_diario = dados['PrecoAcao'].pct_change()
    media_retorno = retorno_diario.mean()
    volatilidade_retorno = retorno_diario.std()
    return media_retorno, volatilidade_retorno


def simulacao_monte_carlo(preco_atual, media_retorno, volatilidade_retorno, dias_simulacoes, num_simulacoes):
    """Realiza a simulação de Monte Carlo."""
    simulacoes = np.zeros((num_simulacoes, dias_simulacoes))

    for i in range(num_simulacoes):
        preco = np.zeros(dias_simulacoes)
        preco[0] = preco_atual

        for dia in range(1, dias_simulacoes):
            retorno_diario = np.random.normal(media_retorno, volatilidade_retorno)
            preco[dia] = preco[dia - 1] * (1 + retorno_diario)

        simulacoes[i, :] = preco

    return simulacoes


def visualizar_simulacoes(simulacoes):
    """Visualiza as simulações de Monte Carlo."""
    plt.figure(figsize=(10, 6))
    plt.plot(simulacoes.T, color='blue', alpha=0.1)
    plt.title('Simulação Monte Carlo para Preço Futuro da Ação')
    plt.xlabel('Dias')
    plt.ylabel('Preço da Ação')
    plt.show()


if __name__ == "__main__":
    nome_arquivo = 'precoAcoes.xlsx'
    planilha = 'precoAcoes'

    dados = carregar_dados(nome_arquivo, planilha)
    media_retorno, volatilidade_retorno = calcular_estatisticas_retorno(dados)

    preco_atual = dados['PrecoAcao'].iloc[-1]
    dias_simulacoes = 252
    num_simulacoes = 1000

    simulacoes = simulacao_monte_carlo(preco_atual, media_retorno, volatilidade_retorno, dias_simulacoes, num_simulacoes)

    # Calcular a média para cada dia
    media_por_dia = np.mean(simulacoes, axis=0)

    # Exibir as médias para cada dia com o número do dia
    print("Média para cada dia:")
    for dia, media in enumerate(media_por_dia, start=1):
        print(f"Dia {dia}: {media}")

    visualizar_simulacoes(simulacoes)
