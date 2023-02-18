from datetime import date
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import Reference

from classes import GerenciadorPlanilhas, LeitorArquivos, PropriedadeSeries

acao = input("Qual o código da Ação que você quer processar?").upper()
# acao = "BIDI4"

try:
    leitor = LeitorArquivos('./dados/')
    leitor.processa_arquivo(acao)
        
    gerenciador = GerenciadorPlanilhas()
    plan_dados = gerenciador.adiciona_planilha('dados')

    gerenciador.adiciona_linha(["DATA", "COTACAO", "BANDA INFERIOR", "BANDA SUPERIOR"])


    indice = 2

    for linha in leitor.dados:
        ano_mes_dia = linha[0].split(" ")[0]
        data = date(
            year = int(ano_mes_dia.split('-')[0]),
            month = int(ano_mes_dia.split('-')[1]), 
            day = int(ano_mes_dia.split('-')[2]),
        )
        
        formula_limite_inferior = f'=AVERAGE(B{indice}:B{indice+19}) - 2*stdev(B{indice}:B{indice+19})'
        formula_limite_superior = f'=AVERAGE(B{indice}:B{indice+19}) + 2*stdev(B{indice}:B{indice+19})'  
        
        gerenciador.atualiza_celula( f'A{indice}', data )
        gerenciador.atualiza_celula( f'B{indice}', float(linha[1]) )
        gerenciador.atualiza_celula( f'C{indice}', formula_limite_inferior )
        gerenciador.atualiza_celula( f'D{indice}', formula_limite_superior )
        
        
        indice += 1


    gerenciador.adiciona_planilha('Grafico')
    gerenciador.mescla_celulas("A1","T2")
    gerenciador.aplica_estilos(
        celula = 'A1',
        estilos = [
            ('font', Font(b=True, sz=18, color="FFFFFF")),
            ('fill', PatternFill('solid', fgColor='07838f')),
            ('alignment', Alignment(vertical='center', horizontal='center')),
        ]
    )

    gerenciador.atualiza_celula('A1', 'Histórico de Cotações')

    referencia_cotacoes = Reference(plan_dados, min_col=2, min_row=2, max_col=4, max_row=indice)
    referencia_datas = Reference(plan_dados, min_col=1, min_row=2, max_col=1, max_row=indice) 

    gerenciador.grafico_linha(
        celula = 'A3',
        comprimento = 33.87,
        altura = 14.82,
        titulo = f'Cotações - {acao}',
        titulo_x = 'Data Cotação',
        titulo_y = 'Valor Cotação',
        referencia_x = referencia_cotacoes,
        referencia_y = referencia_datas,
        propriedades_grafico = [
            PropriedadeSeries(espessura=0, cor='0a55ab'),
            PropriedadeSeries(espessura=0, cor='a61508'),
            PropriedadeSeries(espessura=0, cor='12a154'),
        ]
    )

    # gerenciador.mescla_celulas('I32','L35')
    # gerenciador.adiciona_imagem('I32', caminho)

    gerenciador.salva_arquivo('./saida/CotacaoRef.xlsx')


except AttributeError:
    print('Atributo inexistente!')

except ValueError:
    print('Formato de dados incorreto. Favor, verificar!')
    
except FileNotFoundError:
    print('Arquivo não encontrado!')
    
except Exception as e:
    print(f'Ocorreu um erro: {e}')