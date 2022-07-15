# Conversor de horários de aulas da UEM

Eu fiz esse scriptzinho pra não ter que ficar fazendo manualmente o embelezamento dos horários que a UEM manda, já que tô fazendo pra alguns amigos também, achei que poderia ser útil pra mais alguém então tô compartilhando aqui no GitHub também.

## Como eu uso isso?
1. Renomeia o PDF do seu horário para "horario.pdf"
2. Coloca ele na mesma pasta que o ```main.py```
3. Rode o script

Se você nunca rodou um script de Python antes, é só instalar o Python, colocar ele nas variáveis de ambiente do seu computador, instalar os módulos necessários e rodar.
A sequência de comandos no Prompt de Comando (Windows) ou no terminal (linux) é mais ou menos a seguinte:

```cd "caminho/para/a/pasta"``` <-- Pasta onde os arquivos estão

```pip install requirements.txt```

```python main.py```

Se tudo deu certo, um arquivo chamado "Horario.xlsx" deve ter aparecido, se aconteceu algum erro, existem diversos tutoriais na internet para te ajudar.

PS: Alguns cursos tem uma terceira folha de matérias modulares. Eu ainda não implementei essas daí então elas não vão aparecer.
