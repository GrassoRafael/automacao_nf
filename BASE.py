import pyautogui
import time
import datetime
import PySimpleGUI as sg
from sys import exit
import pandas as pd

# Importar Base de Agenda
agenda = pd.read_excel('D:/Github/automacao_nf')

# Layout
sg.theme('reddit')
layout = [
    [sg.Image(filename='LOGO_U_REDUZ.png'), sg.Text('Empresa'), sg.Push()],
    [sg.Text('Insira as informações para nota fiscal'), sg.Push()],
    [sg.Text('Cod Cliente'), sg.Input(key='cod_cliente')],
    [sg.Text('Peso'), sg.Input(key='peso')],
    [sg.Text('Cliente:', key='nome_cliente'), sg.Push()],
    [sg.Button('Buscar')],
    [sg.Text('Data da Agenda'), sg.Input(key='data_agenda')],
    [sg.Text('Preço do Frete'), sg.Input(key='preco_frete')],
    [sg.Text('Próx. Ticket'), sg.Input(key='ticket')],
    [sg.Text('Próx. NF'), sg.Input(key='nf')],
    [sg.Text('Cod. Motorista'), sg.Combo(['512', '365'], key='cod_motorista', size=(11, 1)), sg.Stretch(),
     sg.Text('Placa Motorista'), sg.Combo(['ABC1234, CBA4321'], key='placa_motorista', size=(11, 1))],
    [sg.Text('Nome Motorista'), sg.Combo(['Nome1', 'Nome2'], key='nome_mot')],
    [sg.Text('Fornecedor     '), sg.Combo(['12345'], key='fornecedor', size=(11, 1)), sg.Stretch(),
     sg.Text('Cod Transação'), sg.Combo(['5102', '6102', '2102'], key='cod_transacao', size=(11, 1))],
    [sg.Text('Produto          '), sg.Combo(['ABC1234565'], key='produto', size=(11, 1)), sg.Stretch(),
     sg.Text('Centro de Custo'), sg.Combo(['1234'], key='cc', size=(11, 1))],
    [sg.Text('Tipo de Frete  '), sg.Combo(['F', 'C'], key='tipo_frete', size=(11, 1)), sg.Stretch(),
     sg.Text('Embutido?'), sg.Combo(['N', 'S'], key='f_embutido', size=(11, 1))],
    [sg.Button('Enviar'), sg.Button('Cancelar')]
]

# Janela
janela = sg.Window('Emissor de Nota Fiscal', layout, element_justification='r')

# Ler os Eventos
while True:
    eventos, valores = janela.read()
    if eventos == sg.WIN_CLOSED or eventos == 'Cancelar':
        exit()
    elif eventos == 'Buscar':
        cod_digitado = valores['cod_cliente']
        try:
            nome_cliente = agenda.loc[agenda['cod'] == int(cod_digitado), 'cliente'].values[0]
            janela['nome_cliente'].update(f'Cliente:{nome_cliente}')

            peso = valores['peso']
            frete_cliente = agenda.loc[agenda['cod'] == int(cod_digitado), 'frete_ton'].values[0]
            janela['preco_frete'].update(f'{float(frete_cliente) * float(peso):.2F}')

            # Puxar e atualizar data da Agenda
            data_agenda = valores['data_agenda']
            agenda['data'] = pd.to_datetime(agenda['data'])
            agenda['data'] = agenda['data'].dt.strftime('%d/%m/%Y')
            agenda_cliente = agenda.loc[agenda['cod'] == int(cod_digitado), 'data'].values[0]
            janela['data_agenda'].update(f'{agenda_cliente}') # Formatar como data

            # Atualizar Ticket
            prox_ticket = agenda.loc[agenda['cod'] == int(cod_digitado), 'ticket'].values[0]
            janela['ticket'].update(f'{prox_ticket}')

            # Atualizar NF
            prox_nf = agenda.loc[agenda['cod'] == int(cod_digitado), 'nf'].values[0]
            janela['nf'].update(f'{prox_nf}')


        except IndexError:
            sg.popup_error('Código não encontrado. Verifique o código digitado')

    if eventos == 'Enviar':
        break

janela.close()

# Variáveis
cod_cliente = valores['cod_cliente']
data_agenda = valores['data_agenda']
fornecedor = valores['fornecedor']
cod_mot = valores['cod_motorista']
placa = valores['placa_motorista']
produto = valores['produto']
cod_trans = valores['cod_transacao']
centro_custo = valores['cc']
p_frete = valores['preco_frete']
tipo_frete = valores['tipo_frete']
frete_emb = valores['f_embutido']
nome_mot = valores['nome_mot']
ticket = str(agenda['ticket'].values[0])
nf = str(agenda['nf'].values[0])

time.sleep(4)
pyautogui.PAUSE = 0.3

# Acessar agenda
pyautogui.click(439, 45)
pyautogui.click(448, 73)
pyautogui.click(722, 69)

# Inserir informações na agenda
time.sleep(6)

# Cliente
pyautogui.click(369, 168)
pyautogui.write(cod_cliente) # Código do cliente
pyautogui.press('tab')
pyautogui.write(data_agenda)
pyautogui.press('tab')
pyautogui.write('1') # Número da carga no mesmo dia do cliente
pyautogui.press('tab')
time.sleep(0.5)

# Sinaleiro
pyautogui.click(188, 261)
pyautogui.moveTo(306, 400)
pyautogui.click(clicks=2)
pyautogui.press('backspace')
pyautogui.press('1')
pyautogui.press('tab')
pyautogui.press('enter')
time.sleep(0.5)
pyautogui.click(248, 266)

# Carga
pyautogui.moveTo(1085, 517)
pyautogui.click(clicks=2, interval=0.25)
pyautogui.click(813, 317)
pyautogui.click(clicks=2, interval=0.25)
pyautogui.press('backspace')
pyautogui.write(peso) # Peso do caminhão
pyautogui.press('up')
pyautogui.click(1183, 136)
time.sleep(0.5)
pyautogui.click(308, 260)

# Frete
time.sleep(0.5)
pyautogui.click(381, 366)
time.sleep(0.3)
pyautogui.click(clicks=2, interval=0.25)
pyautogui.press('tab')
pyautogui.write(peso) # Peso do Caminhão
pyautogui.moveTo(367, 458)
pyautogui.click(clicks=2, interval=0.25)
pyautogui.press('backspace')
pyautogui.write(cod_mot) # Código do motorista
pyautogui.press('tab')
pyautogui.write(placa) # Placa do caminhão
pyautogui.press('tab')
pyautogui.press('enter')

# Acessar Ticket
pyautogui.click(118, 35) # Acesso mercado
pyautogui.click(315, 211)
pyautogui.click(515, 216)
time.sleep(7)

# Inserir informações no ticket
data_atual = datetime.datetime.now()
data_form = data_atual.strftime('%d/%m/%Y')

pyautogui.doubleClick(254, 168)
pyautogui.press('backspace')
pyautogui.write(data_form) # Data do lançamento
pyautogui.press('tab')
pyautogui.write(ticket) # Número sequencial do ticket
pyautogui.press('tab')
pyautogui.write('0')
pyautogui.press('tab')
pyautogui.write(placa) # Placa do veículo
pyautogui.press('tab')
pyautogui.write(fornecedor) # Fornecedor
pyautogui.press('tab')
pyautogui.write(nome_mot) # Nome do motorista
pyautogui.press('tab')
pyautogui.write(produto)
pyautogui.press('tab')
pyautogui.write(peso) # Peso do caminhão
pyautogui.press('tab')
pyautogui.write(peso) # Peso do caminhão
pyautogui.press('up')
pyautogui.click(1205, 138)
time.sleep(2)
pyautogui.press('enter')
time.sleep(4)
pyautogui.press('enter')
time.sleep(2)

# Acessar as notas manuais
pyautogui.click(118, 35) # Acesso mercado
pyautogui.click(353, 127)
pyautogui.click(632, 149)
pyautogui.click(871, 177)

time.sleep(10)

# Inserir informações para nota - Cabeçalho
pyautogui.press(['tab', 'tab', 'tab'], interval=0.5)
time.sleep(2)
pyautogui.press('enter') # Pressionar a tela pop up
pyautogui.press('tab')

pyautogui.write(cod_cliente) # Código do cliente
pyautogui.press('tab')
pyautogui.write(cod_trans) # Transação Prod. 5102 Dentro do Estado - 6102 Transação Prod Fora do estado
pyautogui.click(1113, 105)
time.sleep(6)

# Informações do produto
pyautogui.click(226, 376)
pyautogui.write(produto) # Código do produto
pyautogui.press('tab', presses=3)
pyautogui.write(peso) # Peso da carga
pyautogui.press('tab')
pyautogui.write(peso) # Peso da Carga
pyautogui.press('tab')
pyautogui.write(peso) # Peso da Carga
pyautogui.press('tab', presses=4, interval=0.25)
pyautogui.write('80') # Cta. Financeiro
pyautogui.press('tab')
pyautogui.write(centro_custo) # Centro de Custo
pyautogui.press('up')
pyautogui.click(260, 323)
time.sleep(2)

# Parcelas (Nota fiscal)
pyautogui.tripleClick(1234, 783)
time.sleep(1)
pyautogui.tripleClick(1077, 375)
time.sleep(0.5)
pyautogui.press('backspace')
pyautogui.write('25') # Forma de pagamento
pyautogui.press('up')
pyautogui.click(108, 237)
time.sleep(4)

# Diversos (Nota Fiscal)
pyautogui.click(334, 211)
pyautogui.write(cod_mot) # Cod Motorista
pyautogui.press('tab', presses=2, interval=0.25)
pyautogui.write(placa) # Placa Caminhão
pyautogui.press('tab', presses=6)
pesokg = float(peso) * 1000 # Variável para armazenar o cálculo de peso de TON para KG
pyautogui.write(str(int(pesokg))) # Multiplicar peso por 1000, peso em kg
pyautogui.press('tab')
pyautogui.write('8') # Código da Embalagem, granel (8)
pyautogui.press('tab', presses=5)
pyautogui.write('75') # Operação Isenta de ICMS
pyautogui.click(1313, 143)
pyautogui.click(584, 532)
time.sleep(0.5)
pyautogui.click(249, 237)

# Valores (Nota Fiscal)
pyautogui.click(750, 684)
pyautogui.write(p_frete) # Valor do frete
pyautogui.doubleClick(902, 512, interval=0.25)
pyautogui.press('backspace')
pyautogui.write(tipo_frete) # Tipo de frete, CIF (C) ou FOB (F)
pyautogui.click(1214, 284)

# Finalização da nota
pyautogui.click(1105, 199)
pyautogui.write(frete_emb) # Frete embutido
pyautogui.press('tab')
pyautogui.click(1074, 494)
pyautogui.click(1119, 136) # Confirma o fechamento da nota
pyautogui.press('enter')
time.sleep(5)
pyautogui.press('right')
pyautogui.press('enter')
time.sleep(4)
pyautogui.press('enter')

# Acessar recibo de frete
pyautogui.click(108, 42)
pyautogui.click(290, 210)
pyautogui.click(623, 354)
time.sleep(5)

# Preencher recibo
pyautogui.click(250, 197)
pyautogui.write(nf) # Número NF
pyautogui.press('tab')
pyautogui.write('0')
pyautogui.press('tab')
time.sleep(1.5)
pyautogui.write('5') # Tipo de veículo
pyautogui.press('tab')
pyautogui.write(cod_mot) # Transportadora
pyautogui.press('tab')
pyautogui.write(p_frete) # Preço do frete
pyautogui.press('tab')
pyautogui.click(1189, 140)
time.sleep(1.5)
pyautogui.click(264, 594)
time.sleep(9)

# Salvar Recibo de Frete
pyautogui.click(547, 43)
time.sleep(2)
pyautogui.click(1132, 699)
pyautogui.moveTo(1132, 830)
pyautogui.click(clicks=4, interval=0.5)
time.sleep(1)
pyautogui.click(974, 831)
pyautogui.doubleClick(948, 665)
pyautogui.press('backspace')
pyautogui.write('Recibo de frete ' + nome_cliente) # Recibo + nome do cliente
pyautogui.click(1176, 662)
pyautogui.click(1231, 8)

# Acessar boleto
pyautogui.click(114, 42)
pyautogui.click(246, 120)
pyautogui.click(541, 147)
pyautogui.click(1111, 552)
pyautogui.click(1299, 552)
time.sleep(1.5)
pyautogui.press('tab')
pyautogui.write(nf) # Número da Nota Fiscal
pyautogui.press('tab')
pyautogui.write(nf) # Número da Nota Fiscal 2
pyautogui.click(143, 340)
pyautogui.click(316, 338)
pyautogui.write('001') # Número do Banco do Brasil
pyautogui.click(1222, 135)
pyautogui.click(1227, 195)
pyautogui.click(1217, 166)
pyautogui.press('enter')
time.sleep(0.5)

# Salvar boleto
pyautogui.click(680, 494)
pyautogui.write('17') # Número da Carteira Ullmann
pyautogui.click(1262, 444)
time.sleep(4)
pyautogui.click(1143, 62)
time.sleep(6)
pyautogui.click(547, 43)
time.sleep(3)
pyautogui.click(1132, 699)
pyautogui.moveTo(1132, 830)
pyautogui.click(clicks=4, interval=0.5)
time.sleep(1)
pyautogui.click(974, 831)
pyautogui.doubleClick(948, 665)
pyautogui.press('backspace')
pyautogui.write(nf + ' - ' + nome_cliente) # Nome e nota para boleto
pyautogui.click(1176, 662)
pyautogui.click(1231, 8)
time.sleep(1)
pyautogui.press('enter')
