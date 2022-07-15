# Conversor de horários de aulas da UEM

Eu fiz esse scriptzinho pra não ter que ficar fazendo manualmente o embelezamento dos horários que a UEM manda, já que tô fazendo pra alguns amigos também, achei que poderia ser útil pra mais alguém então tô compartilhando aqui no GitHub também.

Depois de feito, fica muito mais tranquilo formatar e editar.

NOTA: Sempre cheque se todos os dados das aulas estão certos. Não posso garantir que sempre vai sair certo.

Torna isso:

![antigo](https://user-images.githubusercontent.com/88753590/179155178-8413fa2c-0598-4a7e-8147-1c68dcf87075.PNG)


Nisso:

![novo](https://user-images.githubusercontent.com/88753590/179155364-42ad58b2-25c2-4bf8-a80a-eae133f63a54.PNG)

(o segundo semestre aparece logo depois)

## Como eu uso isso?
1. Renomeia o PDF do seu horário para "horario.pdf"
2. Coloca ele na mesma pasta que o ```main.py```
3. Rode o script

Se você nunca rodou um script de Python antes, é só instalar o Python, colocar ele nas variáveis de ambiente do seu computador, instalar os módulos necessários e rodar.
A sequência de comandos no Prompt de Comando (Windows) ou no terminal (linux) é mais ou menos a seguinte:

```cd "caminho/para/a/pasta"``` <-- Pasta onde os arquivos estão

```pip install -r requirements.txt```

```python main.py```

Se tudo deu certo, um arquivo chamado "Horario.xlsx" deve ter aparecido, se aconteceu algum erro, existem diversos tutoriais na internet para te ajudar.

PS: Alguns cursos tem uma terceira folha de matérias modulares. Eu ainda não implementei essas daí então elas não vão aparecer.
