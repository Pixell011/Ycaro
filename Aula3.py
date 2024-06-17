#if / elif / else
#se/ se não se / se não

from os import path


entrada = input ('Você quer "entrar"ou "sair"?')
if entrada == 'entrar':
 print ('Você entrou no sistema')
elif entrada == "sair":
 print ('Você saiu do sistema')
else:
 print('Você não digitou nem entrar nem sair')
 