from datetime import datetime, timedelta 


ontem = datetime.now() - timedelta(1) - timedelta(1)
ontem_formatado_preencher = str(ontem.strftime('%d/%m/%Y'))
ontem_formatado = str(ontem.strftime('%d%m%Y'))

dia = int(ontem_formatado[0:2])
mes = int(ontem_formatado[2:4])
ano = int(ontem_formatado[4:8])

ontem_formatado = datetime(ano, mes, dia, 0, 0)

print(ontem)
print(ontem_formatado_preencher)
print(ontem_formatado)

data = '24/01/2024'

dia_digitado = int(data[0:2])
mes_digitado = int(data[3:5])
ano_digitado = int(data[6:10])
dia_formatado = datetime(ano_digitado, mes_digitado, dia_digitado, 0, 0)
print(dia_formatado)