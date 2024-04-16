import datetime
import locale

# Definir a localização para português do Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Obter o nome do mês atual
nome_mes_atual = datetime.datetime.now().strftime('%B')

ano_atual = datetime.datetime.now().strftime('%Y')

print(ano_atual)