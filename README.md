# Dicionário de Python para Automação

Guia prático de **Python, regex e automação**, com exemplos tirados de um projeto real.

## Sobre este guia

Este material reúne o que aprendi ao automatizar uma tarefa repetitiva do meu trabalho: ler documentos Word e preencher um formulário web com os dados deles.

Cada conceito é explicado em linguagem simples, sempre respondendo a três perguntas:

- **Para que serve?**
- **Como funciona?**
- **Onde eu uso?**

**Para quem é:** quem está começando em Python ou quer automatizar tarefas do dia a dia. Não é preciso saber programar.

**Como usar:** leia na ordem ou vá direto ao tema pelo sumário. No final há 15 exercícios com gabarito para praticar.

---

## Sumário

1. [O código que eu uso de exemplo](#1-o-código-que-eu-uso-de-exemplo)
2. [Regex: o que é](#2-regex-o-que-é)
3. [Regex: os símbolos](#3-regex-os-símbolos)
4. [Regex: os comandos do Python](#4-regex-os-comandos-do-python)
5. [Python: o básico de tudo](#5-python-o-básico-de-tudo)
6. [Python: tomando decisões (if, elif, else)](#6-python-tomando-decisões-if-elif-else)
7. [Python: repetindo coisas (for, while)](#7-python-repetindo-coisas-for-while)
8. [Python: funções (def, return)](#8-python-funções-def-return)
9. [Python: lidando com erros (try, except)](#9-python-lidando-com-erros-try-except)
10. [Python: outras coisas que aparecem muito](#10-python-outras-coisas-que-aparecem-muito)
11. [Bibliotecas que eu uso](#11-bibliotecas-que-eu-uso)
12. [Selenium: mexendo no navegador](#12-selenium-mexendo-no-navegador)
13. [Outras formas de automatizar](#13-outras-formas-de-automatizar)
14. [Exercícios](#14-exercícios)
15. [Gabarito](#15-gabarito)

---

## 1. O código que eu uso de exemplo

No meu trabalho, eu recebia um documento Word com dados de um pagamento (placa, datas, valor, dados do banco) e tinha que **copiar tudo, campo por campo**, para um formulário num site.

Então eu quis fazer um robô em Python que fizesse isso por mim. Ele faz 3 coisas:

1. **Abre o Word** e lê tudo o que está escrito.
2. **Procura cada informação** (a placa, a data, o valor...) e guarda numa espécie de armário de gavetas.
3. **Abre o navegador** e digita cada informação no lugar certo do formulário.

Ele **não aperta o botão de enviar**. Essa parte fica comigo, porque eu confiro tudo antes.

### A primeira versão não deu certo

A primeira versão usava o **Playwright** e não preenchia tudo. Alguns campos ficavam em branco ou com o valor errado. Os problemas eram:

- **Esperas fixas:** o robô esperava 4 segundos e já começava a digitar. Se o site demorava mais, os campos ainda não existiam.
- **Datas e bolinhas (rádio) difíceis:** a data do formulário tem 3 quadradinhos separados (dia, mês e ano), e o robô tentava digitar tudo num lugar só.

### A versão que funciona

Com a ajuda de dois amigos, o código foi refeito com o **Selenium**, e essa versão funciona. O que mudou:

| Problema da primeira versão | Como a versão nova resolve |
| --- | --- |
| Espera fixa de 4 segundos | Espera **até** o campo aparecer (`WebDriverWait`), no máximo 8 segundos |
| Data digitada num lugar só | Acha os 3 quadradinhos e digita dia, mês e ano, cada um no seu |
| Erro no meio travava tudo | Mostra uma mensagem clara de qual campo falhou e salva tudo num `log.txt` |

A lição que eu tirei: **o robô precisa esperar o site carregar antes de agir**, igual uma pessoa faria.

Sempre que eu falar "no meu código", é dessa versão que funciona.

---

## 2. Regex: o que é

Regex é um jeito de dizer para o computador **qual é a cara** do que você está procurando.

Pensa assim: se eu te peço "procura uma data nesse papel", você bate o olho e acha, porque você sabe que data tem essa cara:

```text
__/__/____
```

O computador não sabe disso. Então eu preciso **desenhar essa cara** para ele, usando símbolos:

```text
\d\d/\d\d/\d\d\d\d
```

O `\d` quer dizer "aqui vai um número". Então isso aí é: número, número, barra, número, número, barra, número, número, número, número.

Ou seja: `07/11/2026`.

**É só isso.** Regex é desenhar a cara da coisa com símbolos. O resto é aprender quais símbolos existem.

---

## 3. Regex: os símbolos

Não precisa decorar. Volta aqui sempre que travar.

### 3.1 Símbolos que valem UMA coisa

| Símbolo | O que eu leio quando vejo ele | Exemplo |
| --- | --- | --- |
| `abc` | "exatamente essas letras" | `PLACA` acha a palavra PLACA |
| `\d` | "um número" | `\d` acha o `7` |
| `\D` | "qualquer coisa que **não** é número" | em `12.345-6` acha o ponto e o traço |
| `\s` | "um espaço" | o espaço no meio de `DATA DA` |
| `\S` | "qualquer coisa que **não** é espaço" | letras, números, sinais |
| `.` | "**qualquer** coisa" (é um coringa) | `N.MERO` acha NÚMERO e NUMERO |
| `[abc]` | "uma dessas da lista" | `[\d.]` = um número **ou** um ponto |
| `[^abc]` | "qualquer uma, **menos** essas" | `[^\r\n]` = qualquer coisa, menos pular linha |
| `\r\n` | "o Enter" (a quebra de linha) | aparece dentro de `[^\r\n]` |
| `\` | "não é símbolo, é o sinal mesmo" | `\.` = um ponto de verdade |

**No meu código:**

```python
r"^\s*N.MERO\s+DO\s+SINISTRO\s*:\s*(\S+)"
```

Eu usei o `.` no lugar do Ú porque às vezes o Word vem com o acento estragado. Com o coringa, tanto faz.

### 3.2 Símbolos que dizem QUANTAS VEZES

Esses sempre vêm **depois** de outro símbolo. Eles dizem quantas vezes aquele símbolo pode aparecer.

| Símbolo | O que eu leio | Exemplo |
| --- | --- | --- |
| `+` | "1 ou mais vezes" | `\d+` = um número de qualquer tamanho: `7`, `45821` |
| `*` | "0 ou mais vezes" (pode nem ter) | `\s*` = com espaço **ou sem** espaço |
| `?` | "pode ter ou não" | `R?` = pode ter a letra R ou não |
| `{4}` | "exatamente 4 vezes" | `\d{4}` = `2026` |
| `{1,2}` | "de 1 a 2 vezes" | `\d{1,2}` = `7` ou `07` |
| `+?` ou `*?` | "o **mínimo** possível" (preguiçoso) | para no primeiro lugar que der certo |

**Por que eu uso tanto `\s*`?** Porque cada pessoa digita de um jeito. Uma escreve `PLACA:ABC`, outra escreve `PLACA:   ABC`. Com `\s*`, os dois funcionam.

### 3.3 Símbolos que marcam LUGAR

| Símbolo | O que eu leio | Exemplo |
| --- | --- | --- |
| `^` | "no começo da linha" | `^PLACA` = PLACA logo no começo |
| `$` | "no fim da linha" | `\d$` = número no final |
| `\b` | "na borda de uma palavra" | `\bSPLIT\b` acha SPLIT, mas não acha SPLITRISK |

### 3.4 Símbolos que ORGANIZAM

| Símbolo | O que eu leio | Exemplo |
| --- | --- | --- |
| `( )` | "**guarda** isso numa gaveta" | `PLACA:\s*(.+)` guarda só a placa |
| `(?: )` | "junta, mas **não** guarda" | `(?:CNPJ\|CPF)` |
| `\|` | "ou" | `ROUBO\|FURTO` = um ou outro |

**A gaveta `( )` é o mais importante de tudo.** Olha:

No Word está escrito:

```text
PLACA: ABC1D23
```

E eu uso este regex:

```text
^\s*PLACA\s*:\s*([^\r\n]+)
```

Lendo pedaço por pedaço:

| Pedaço | O que eu leio |
| --- | --- |
| `^` | no começo da linha... |
| `\s*` | ...pode ter espaço ou não... |
| `PLACA` | ...vem escrito PLACA... |
| `\s*:\s*` | ...vem os dois-pontos (com ou sem espaço)... |
| `(` | ...**abre a gaveta**... |
| `[^\r\n]+` | ...pega tudo até o fim da linha... |
| `)` | ...**fecha a gaveta**. |

Resultado: na gaveta fica **só** `ABC1D23`. A palavra PLACA fica de fora, porque está fora da gaveta.

**Se você entendeu esse, entendeu quase todos do meu código.** Os de segurado, terceiro, apólice e PIX são iguaizinhos. Só muda a palavra.

### 3.5 O `r` antes das aspas

Em regex eu **sempre** escrevo um `r` antes das aspas:

```python
r"\d"
```

Esse `r` diz para o Python: "não mexe nas minhas barras". Sem ele, o Python pode achar que a barra é outra coisa e bagunçar tudo.

---

## 4. Regex: os comandos do Python

Para usar o regex no Python, primeiro eu escrevo `import re` no topo do arquivo. Depois eu uso estes comandos:

| Comando | O que ele faz | Exemplo |
| --- | --- | --- |
| `re.search(regex, texto)` | Procura e me traz o **primeiro** que achar | achar a placa |
| `re.findall(regex, texto)` | Procura e me traz **todos** numa lista | achar todas as datas |
| `re.sub(regex, troca, texto)` | Procura e **troca** por outra coisa | `re.sub(r"\D", "", "12.345-6")` vira `123456` |
| `re.fullmatch(regex, texto)` | Pergunta: "o texto **inteiro** tem essa cara?" | conferir se a data está certa |
| `.group(1)` | Abre a gaveta 1 e me mostra o que tem dentro | `achado.group(1)` |
| `.groups()` | Abre **todas** as gavetas de uma vez | `banco.groups()` |

E dois ajustes que eu uso muito:

- `re.IGNORECASE` → "tanto faz maiúscula ou minúscula".
- `re.MULTILINE` → "o `^` vale para o começo de **cada** linha".

**No meu código**, eu criei uma função para não repetir isso toda hora:

```python
def buscar(padrao, default=""):
    achado = re.search(padrao, blob, re.IGNORECASE | re.MULTILINE)
    return achado.group(1).strip() if achado else default
```

Lendo em português: "procura. Se achou, me dá o que está na gaveta 1, sem espaços sobrando. Se não achou, me dá vazio."

---

## 5. Python: o básico de tudo

### 5.1 Regras que valem para o código inteiro

1. **O Python lê de cima para baixo**, uma linha de cada vez. Igual a gente lendo um livro.
2. **Os espaços no começo da linha importam muito.** Eles mostram o que está "dentro" de quê. Use sempre 4 espaços.
3. **Os dois-pontos `:` abrem um bloco.** Tudo o que vem embaixo, com 4 espaços, faz parte dele.
4. **`#` é comentário.** O Python ignora. Serve para eu deixar recadinhos no código.
5. **Maiúscula é diferente de minúscula.** `placa` e `Placa` são duas coisas diferentes.

```python
if valor:                 # os dois-pontos abrem o bloco
    print("tem valor")    # 4 espaços: está DENTRO do if
print("sempre aparece")   # sem espaço: está FORA do if
```

### 5.2 Variável

Variável é uma **caixinha com etiqueta** onde eu guardo uma coisa.

```python
placa = "ABC1D23"
ESPERA = 8
```

Agora, toda vez que eu escrever `placa`, o Python sabe que é `"ABC1D23"`.

Quando eu escrevo o nome **todo em maiúsculo** (`ESPERA`), é um combinado entre programadores: "isso aqui não muda".

### 5.3 Tipos de dados

Cada coisa que eu guardo tem um tipo:

| Tipo | O que é | Exemplo |
| --- | --- | --- |
| `str` | Texto (sempre entre aspas) | `"ABC1D23"` |
| `int` | Número inteiro | `8` |
| `float` | Número com vírgula (no Python é **ponto**) | `1500.50` |
| `bool` | Verdadeiro ou falso | `True`, `False` |
| `None` | Nada. Vazio. | quando o `re.search` não acha nada |
| `list` | Uma lista de coisas, em ordem | `["Oficina", "Acordo"]` |
| `tuple` | Uma lista que não pode mudar | `("dia", "mês", "ano")` |
| `dict` | Um armário de gavetas com etiqueta | `{"placa": "ABC1D23"}` |

**Pegadinha que me pegou:** `"1.500,00"` é texto, não número. E o Python não entende vírgula como decimal. Para virar número, eu faço assim:

```python
valor = float("1.500,00".replace(".", "").replace(",", "."))
# tira o ponto: "1500,00" → troca a vírgula por ponto: "1500.00" → vira número: 1500.0
```

### 5.4 Lista

Lista é uma **fila de coisas**, uma atrás da outra.

```python
tipos = ["Oficina", "Acordo Segurado", "Acordo Terceiro"]
tipos[0]   # "Oficina"   ← a contagem começa no ZERO
tipos[1]   # "Acordo Segurado"
```

### 5.5 Dicionário

Dicionário é um **armário de gavetas**. Cada gaveta tem uma etiqueta (a chave) e uma coisa dentro (o valor).

```python
dados = {
    "placa": "ABC1D23",
    "numero_sinistro": "45821",
    "valor_total": "1.500,00",
}

dados["placa"]   # "ABC1D23"
```

**No meu código**, é aqui que eu guardo tudo o que achei no Word. Depois, na hora de preencher o site, eu só abro a gaveta certa: `dados["placa"]`.

### 5.6 Comparando coisas

Quando eu quero perguntar algo para o Python, eu uso estes sinais. A resposta é sempre `True` (sim) ou `False` (não).

| Sinal | Pergunta | Exemplo |
| --- | --- | --- |
| `==` | "é igual?" | `tipo == "Oficina"` |
| `!=` | "é diferente?" | `tipo != "Oficina"` |
| `>` `<` | "é maior? é menor?" | `valor > 1000` |
| `>=` `<=` | "é maior ou igual? menor ou igual?" | `tentativas <= 3` |
| `in` | "está dentro?" | `"OFICINA" in nome` |
| `not in` | "não está dentro?" | `"-" not in conta` |

**Muito cuidado:** um `=` só **guarda** (`x = 5`). Dois `==` **perguntam** (`x == 5`).

### 5.7 E, OU, NÃO

| Palavra | O que quer dizer | Exemplo |
| --- | --- | --- |
| `and` | "as duas coisas têm que ser verdade" | `placa and valor` |
| `or` | "basta uma ser verdade" | `tipo == "Roubo" or tipo == "Furto"` |
| `not` | "o contrário" | `not valor` = "não tem valor" |

**Coisas que o Python considera "falso":** `False`, `None`, `0`, `""` (texto vazio), `[]` (lista vazia), `{}` (dicionário vazio). Todo o resto é "verdadeiro".

Por isso eu consigo escrever `if not valor:` para saber se um campo veio vazio do Word.

---

## 6. Python: tomando decisões (if, elif, else)

### Para que serve

Para o código **escolher um caminho**. É o "se acontecer isso, faz aquilo".

### Como funciona

- `if` = **"se"**. Sempre vem primeiro.
- `elif` = **"senão, se"**. Pode ter quantos quiser, no meio.
- `else` = **"senão"**. Vem por último, e é opcional. Roda quando nada lá em cima deu certo.

O Python testa **de cima para baixo** e entra **só no primeiro** que for verdade. Os outros ele pula.

### Exemplo da vida real

```python
if chovendo:
    levar = "guarda-chuva"
elif sol_forte:
    levar = "protetor solar"
else:
    levar = "nada"
```

### Exemplo do trabalho

```python
if "OFICINA" in nome:
    tipo = "Oficina"
elif "TERCEIRO" in nome:
    tipo = "Acordo Terceiro"
elif "SEGURADO" in nome:
    tipo = "Acordo Segurado"
else:
    tipo = "Não identificado"
```

### No meu código

```python
if not valor:
    log(f"⚠ {rotulo}: sem valor no DOCX; deixado em branco")
    return
```

Em português: "se esse campo veio vazio do Word, anota no caderninho e não digita nada".

### `if` numa linha só

Quando a decisão é bem pequena, dá para escrever numa linha:

```python
texto = so_digitos(valor) if numeros else valor
```

Em português: "se for campo de número, usa só os dígitos. Senão, usa do jeito que está".

### Onde eu uso

Sempre que a resposta **depende** de alguma coisa: qual o tipo de pagamento, se o campo está vazio, se é CPF ou CNPJ.

---

## 7. Python: repetindo coisas (for, while)

### 7.1 `for` — "para cada um"

**Para que serve:** fazer a mesma coisa **para cada item** de uma lista.

**Como funciona:** o `for` pega um item por vez, coloca numa variável e roda o bloco. Depois pega o próximo. Até acabar.

```python
frutas = ["maçã", "banana", "uva"]
for fruta in frutas:
    print(fruta)
# maçã
# banana
# uva
```

**No meu código:**

```python
for tabela in documento.tables:
    for linha in tabela.rows:
        ...
```

Em português: "para cada tabela do Word, e para cada linha dessa tabela, faz isso aqui".

**Ajudantes do `for`:**

| Ajudante | O que faz | Exemplo |
| --- | --- | --- |
| `range(3)` | Conta 0, 1, 2 | `for i in range(3):` |
| `enumerate()` | Me dá o número da posição junto | `for i, item in enumerate(lista, 1):` |
| `.items()` | Percorre etiqueta e valor do dicionário | `for chave, valor in dados.items():` |
| `zip()` | Anda em duas listas ao mesmo tempo | `for a, b in zip(lista1, lista2):` |

### 7.2 `while` — "enquanto"

**Para que serve:** repetir **enquanto** uma coisa for verdade. Uso quando eu não sei quantas vezes vai precisar.

```python
tentativa = 1
while tentativa <= 3:
    print(f"Tentativa {tentativa}")
    tentativa += 1    # o mesmo que: tentativa = tentativa + 1
```

**Cuidado:** se a condição nunca virar falsa, ele roda para sempre. Se acontecer, aperta **Ctrl+C** no terminal para parar.

### 7.3 `break`, `continue` e `pass`

| Palavra | O que faz | Quando eu uso |
| --- | --- | --- |
| `break` | **Para** a repetição na hora | Já achei o que eu queria |
| `continue` | **Pula** esse e vai para o próximo | Esse item não me interessa |
| `pass` | Não faz nada | Guardar lugar para escrever depois |

```python
for arquivo in arquivos:
    if not arquivo.endswith(".docx"):
        continue        # não é Word? pula
    if arquivo == "PARAR.docx":
        break           # chegou aqui? para tudo
    print(arquivo)
```

---

## 8. Python: funções (def, return)

### Para que serve

Função é uma **receita guardada na gaveta**. Eu escrevo uma vez e uso quantas vezes eu quiser.

### Como funciona

1. `def nome():` **guarda** a receita. Nesse momento ela **não roda**.
2. Ela só roda quando eu **chamo pelo nome**: `nome()`.
3. O que vai dentro dos parênteses são os **ingredientes** (parâmetros).
4. `return` é **o prato pronto**: o que a função devolve. Quando chega no `return`, a função acaba.

```python
def so_digitos(texto):
    return re.sub(r"\D", "", texto)

conta = so_digitos("12.345-6")
print(conta)   # 123456
```

### Valor padrão

Dá para deixar um ingrediente "já pronto", caso a pessoa não mande:

```python
def buscar(padrao, default=""):
    ...
```

Se eu não mandar o `default`, ele usa `""` (vazio).

### Onde eu uso

Sempre que eu me pego **copiando e colando** o mesmo código. No meu robô, a função `preencher()` é chamada mais de 10 vezes, uma para cada campo do formulário. Imagina escrever tudo 10 vezes!

---

## 9. Python: lidando com erros (try, except)

### Para que serve

Para um erro **não derrubar o programa inteiro**. É tipo uma rede de proteção.

### Como funciona

| Bloco | Quando ele roda |
| --- | --- |
| `try` | Sempre. É o "tenta fazer isso" |
| `except` | **Só se deu erro** no `try` |
| `else` | **Só se NÃO deu erro** no `try` |
| `finally` | **Sempre**, com erro ou sem erro |

```python
try:
    numero = int("abc")
except ValueError:
    print("Isso não é um número!")
else:
    print("Deu certo:", numero)
finally:
    print("Terminei de tentar.")
```

### No meu código

Eu uso o `finally` para garantir que o navegador **sempre fecha**, mesmo se der erro no meio:

```python
try:
    driver.get(FORM_URL)
    # ... preenche tudo ...
finally:
    driver.quit()
```

### Os erros que mais aparecem (e o que eles querem dizer)

| Erro | Em português |
| --- | --- |
| `SyntaxError` | Escrevi algo errado. Faltou `:`, parêntese ou aspas |
| `IndentationError` | Os espaços do começo da linha estão errados |
| `NameError` | Usei um nome que não existe (ou escrevi errado) |
| `TypeError` | Misturei coisas, tipo somar texto com número |
| `ValueError` | O tipo está certo, mas o valor não serve (`int("abc")`) |
| `KeyError` | Pedi uma gaveta que não existe no dicionário |
| `IndexError` | Pedi uma posição que não existe na lista |
| `AttributeError` | Usei um comando em algo que não tem ele (ex.: `.group()` em `None`) |
| `FileNotFoundError` | O arquivo não está onde eu disse que estava |

**Dica:** quando der erro, lê a **última linha** da mensagem. É lá que o Python diz o que aconteceu.

---

## 10. Python: outras coisas que aparecem muito

### `import` — pegar ferramentas

Antes de cozinhar, eu pego as panelas. O `import` é isso.

| Jeito | O que faz | Exemplo |
| --- | --- | --- |
| `import re` | Pega a caixa inteira | uso `re.search()` |
| `from pathlib import Path` | Pega só uma ferramenta da caixa | uso `Path()` direto |
| `import pandas as pd` | Pega e dá um apelido | uso `pd.read_excel()` |

### `with` — abrir e fechar sozinho

```python
with open("log.txt", "w", encoding="utf-8") as arquivo:
    arquivo.write("Terminei!")
# aqui embaixo o arquivo já fechou sozinho
```

### f-string — texto com variável dentro

```python
nome = "Oficina X"
print(f"Pagamento para {nome}")   # Pagamento para Oficina X
```

O `f` antes das aspas deixa eu colocar variáveis dentro de `{ }`.

### `class` — uma caixa de ferramentas com nome

`class` junta várias funções que trabalham juntas. No meu código, a `class FormularioPipefy` guarda tudo o que o robô sabe fazer no formulário: achar campo, digitar, escolher opção, marcar bolinha.

- `self` = "eu mesma". É como a caixa fala dela mesma.
- `__init__` = o que acontece quando a caixa é criada.

### `lambda` — uma função pequenininha

```python
lambda _: radio.is_selected()
```

É uma função de uma linha só, sem nome. Uso quando preciso de algo rapidinho, só ali.

### Outros que aparecem no meu código

| Coisa | O que faz | Exemplo |
| --- | --- | --- |
| `.strip()` | Tira espaços do começo e do fim | `"  ABC ".strip()` → `"ABC"` |
| `.upper()` | Deixa tudo maiúsculo | `"abc".upper()` → `"ABC"` |
| `.replace(a, b)` | Troca a por b | `"1,5".replace(",", ".")` → `"1.5"` |
| `.join()` | Junta uma lista num texto só | `"\n".join(linhas)` |
| `.split()` | Quebra um texto numa lista | `"a-b".split("-")` → `["a", "b"]` |
| `all()` | "**Todos** são verdade?" | todas as palavras estão no nome? |
| `any()` | "**Algum** é verdade?" | alguma palavra está no texto? |
| `input()` | Espera eu digitar algo no terminal | `input("Aperte ENTER")` |
| `print()` | Mostra na tela | `print("Oi")` |
| `sys.argv` | O que eu digitei quando rodei o script | `sys.argv[1]` = o caminho do Word |
| `if __name__ == "__main__":` | "Só começa se rodarem **este** arquivo" | última linha do código |

---

## 11. Bibliotecas que eu uso

Biblioteca é um **pacote de ferramentas prontas** que alguém já fez. Eu só instalo e uso.

| Biblioteca | Para que eu uso |
| --- | --- |
| `re` | Regex (já vem com o Python) |
| `docx` (python-docx) | Ler o Word |
| `unicodedata` | Tirar acentos (já vem com o Python) |
| `pathlib` | Mexer com nomes de arquivo e pastas (já vem com o Python) |
| `sys` | Ler o que eu digitei no terminal (já vem com o Python) |
| `selenium` | Mexer no navegador |

Para instalar as que não vêm com o Python, eu abro o terminal e digito:

```bash
python -m pip install python-docx selenium
```

---

## 12. Selenium: mexendo no navegador

Selenium é o que faz o Python **mexer no navegador sozinho**: abrir site, clicar, digitar.

| Comando | O que faz |
| --- | --- |
| `webdriver.Chrome()` | Abre o Chrome |
| `driver` | É o "controle remoto" do navegador |
| `driver.get(endereço)` | Entra no site |
| `driver.quit()` | Fecha o navegador |
| `find_element()` | Procura **um** elemento na página (botão, campo...) |
| `find_elements()` | Procura **vários** |
| `By.ID`, `By.XPATH`, `By.CSS_SELECTOR` | Diz **como** procurar |
| `.click()` | Clica |
| `.send_keys("texto")` | Digita |
| `Keys.ENTER`, `Keys.CONTROL` | Aperta teclas especiais |
| `WebDriverWait(driver, 8)` | Espera **até** 8 segundos por alguma coisa |
| `EC.element_to_be_clickable()` | "Espera até dar para clicar" |
| `execute_script()` | Roda um código JavaScript dentro da página |
| `TimeoutException` | O erro de "esperei e não apareceu" |

**Por que eu espero em vez de ir direto?** Porque o site demora para carregar. Se o robô tenta digitar antes do campo aparecer, dá erro.

**No meu código:**

```python
self.wait.until(EC.element_to_be_clickable((By.XPATH, seletor))).click()
```

Em português: "espera até 8 segundos o botão ficar clicável. Quando ficar, clica".

---

## 13. Outras formas de automatizar

O Selenium não é o único jeito. A regra que eu sigo:

> **Tem API? Usa API. É site? Usa navegador. É arquivo? Usa a biblioteca do arquivo. Não tem jeito? Usa mouse e teclado.**

### Navegador

| Ferramenta | O que é, resumindo |
| --- | --- |
| **Selenium** | O que eu uso. Mexe no navegador |
| **Playwright** | Igual ao Selenium, só que mais novo. Espera os campos sozinho e tem um gravador que gera o código enquanto você clica |
| **Requests + BeautifulSoup** | Lê os dados de um site **sem abrir** o navegador. Só serve para ler, não para clicar |

### API (o jeito mais confiável)

| Ferramenta | O que é, resumindo |
| --- | --- |
| **Requests** | Conversa direto com o sistema, sem tela nenhuma. Se o sistema tem API, é o melhor caminho |
| **gspread** | Lê e escreve no Google Sheets |

### Arquivos

| Ferramenta | O que é, resumindo |
| --- | --- |
| **python-docx** | Lê e cria Word |
| **openpyxl** | Lê e edita Excel sem estragar a formatação |
| **pandas** | Faz conta, filtro e cruzamento em planilhas grandes |
| **pdfplumber** | Tira texto e tabelas de PDF |
| **reportlab** | Cria PDF |
| **os / shutil / pathlib** | Cria, move, copia e renomeia arquivos e pastas |

### Mouse e teclado

| Ferramenta | O que é, resumindo |
| --- | --- |
| **PyAutoGUI** | Mexe o mouse e digita como se fosse uma pessoa. Funciona em qualquer programa, mas erra fácil se a tela mudar |
| **pywinauto** | Mexe em programas do Windows |
| **pyperclip** | Copia e cola pela área de transferência |

### E-mail e mensagens

| Ferramenta | O que é, resumindo |
| --- | --- |
| **smtplib** | Manda e-mail |
| **imaplib** | Lê e-mails e baixa anexos |
| **API do WhatsApp Business** | Manda mensagem pelo WhatsApp (é pago) |

### Imagem

| Ferramenta | O que é, resumindo |
| --- | --- |
| **pytesseract** | Lê texto de foto ou documento escaneado |
| **Pillow** | Corta, diminui e converte imagens |

### Agendar

| Ferramenta | O que é, resumindo |
| --- | --- |
| **schedule** | Roda uma tarefa num horário, com o script aberto |
| **Agendador de Tarefas do Windows** | Roda o script sozinho todo dia, mesmo com tudo fechado |

### Sem Python (vale conhecer)

| Ferramenta | O que é, resumindo |
| --- | --- |
| **Power Automate Desktop** | Automação da Microsoft, arrastando blocos. Grátis no Windows |
| **Google Apps Script** | Automação dentro do Google Sheets, Docs e Gmail |
| **n8n / Make / Zapier** | Ligam um sistema no outro sem programar |
| **Tampermonkey** | Roda um script dentro de um site, no próprio navegador |

---

## 14. Exercícios

Cria um arquivo chamado `exercicios.py` e escreve `import re` na primeira linha. Faz um de cada vez e testa.

### Regex

1. Escreve um regex que ache uma placa no formato `ABC1D23` (3 letras, 1 número, 1 letra, 2 números).
2. Do texto `"VALOR: R$ 2.350,90"`, pega só o `2.350,90` usando uma gaveta `( )`.
3. Usa o `re.sub` para transformar o CPF `"123.456.789-00"` em `"12345678900"`.
4. Desse texto, usa o `re.findall` para pegar **todas** as datas:

   ```text
   Ocorrência: 01/11/2026 - Comunicação: 03/11/2026 - Pagamento: 10/11/2026
   ```

5. Do mesmo texto, pega **só** a data de pagamento.
6. Explica com as suas palavras a diferença entre `.+` e `.+?`.

### Python

7. Cria um dicionário com `placa`, `segurado` e `valor`, e mostra cada etiqueta e valor usando `for`.
8. Cria uma função `normalizar(texto)` que tira os acentos e deixa tudo maiúsculo.
9. Cria uma função `detectar_tipo(nome_arquivo)` que devolve `"Oficina"` se o nome tiver "OFICINA" e `"Outro"` se não tiver.
10. Usa `try` / `except` para tentar transformar `"abc"` em número e mostra uma mensagem amigável se der erro.
11. Cria uma função `classificar(valor)` com `if` / `elif` / `else`: até 1.000 é `"baixo"`, até 5.000 é `"médio"`, acima disso é `"alto"`.
12. Com a lista `["a.docx", "b.pdf", "c.docx", "PARAR.docx", "d.docx"]`, usa `for`, `continue` e `break` para mostrar só os `.docx` e parar quando chegar no `PARAR.docx`.
13. Usa `while` para mostrar "Tentativa 1", "Tentativa 2" e "Tentativa 3".
14. Transforma o valor `"1.500,00"` em número e soma 250.

### Desafio

15. Cria um Word com estas 3 linhas:

    ```text
    SEGURADO: MARIA SILVA
    PLACA: ABC1D23
    DATA DA OCORRÊNCIA: 05/10/2026
    ```

    Depois faz um script que lê esse Word e monta um dicionário com os 3 dados.

---

## 15. Gabarito

Tenta fazer antes de olhar! Clica no exercício para ver a resposta.

<details>
<summary>1. Placa</summary>

```python
re.search(r"[A-Z]{3}\d[A-Z]\d{2}", "PLACA: ABC1D23").group()
# 'ABC1D23'
```

`[A-Z]{3}` = 3 letras maiúsculas. `\d` = 1 número. `[A-Z]` = 1 letra. `\d{2}` = 2 números.
</details>

<details>
<summary>2. Valor</summary>

```python
re.search(r"R\$\s*([\d.]+,\d{2})", "VALOR: R$ 2.350,90").group(1)
# '2.350,90'
```

O `\$` tem barra porque o cifrão sozinho é um símbolo do regex (fim da linha).
</details>

<details>
<summary>3. Só os números</summary>

```python
re.sub(r"\D", "", "123.456.789-00")
# '12345678900'
```

"Troca tudo que não é número por nada."
</details>

<details>
<summary>4. Todas as datas</summary>

```python
texto = "Ocorrência: 01/11/2026 - Comunicação: 03/11/2026 - Pagamento: 10/11/2026"
re.findall(r"\d{2}/\d{2}/\d{4}", texto)
# ['01/11/2026', '03/11/2026', '10/11/2026']
```
</details>

<details>
<summary>5. Só a data de pagamento</summary>

```python
re.search(r"Pagamento:\s*(\d{2}/\d{2}/\d{4})", texto).group(1)
# '10/11/2026'
```

O segredo é usar a palavra que vem **logo antes** da data como ponto de partida.
</details>

<details>
<summary>6. <code>.+</code> x <code>.+?</code></summary>

`.+` é **guloso**: come o máximo que conseguir. `.+?` é **preguiçoso**: para assim que der.

```python
re.search(r"<(.+)>", "<a><b>").group(1)   # 'a><b'  (comeu tudo)
re.search(r"<(.+?)>", "<a><b>").group(1)  # 'a'     (parou no primeiro)
```
</details>

<details>
<summary>7. Dicionário com for</summary>

```python
dados = {"placa": "ABC1D23", "segurado": "MARIA SILVA", "valor": "1.500,00"}
for chave, valor in dados.items():
    print(f"{chave}: {valor}")
```
</details>

<details>
<summary>8. normalizar</summary>

```python
import unicodedata

def normalizar(texto):
    texto = unicodedata.normalize("NFKD", texto or "")
    return texto.encode("ascii", "ignore").decode().upper().strip()

normalizar("Colisão")  # 'COLISAO'
```
</details>

<details>
<summary>9. detectar_tipo</summary>

```python
def detectar_tipo(nome_arquivo):
    if "OFICINA" in nome_arquivo.upper():
        return "Oficina"
    return "Outro"
```
</details>

<details>
<summary>10. try / except</summary>

```python
try:
    numero = int("abc")
except ValueError:
    print("Isso não é um número.")
```
</details>

<details>
<summary>11. if / elif / else</summary>

```python
def classificar(valor):
    if valor <= 1000:
        return "baixo"
    elif valor <= 5000:
        return "médio"
    else:
        return "alto"

classificar(800)    # 'baixo'
classificar(3000)   # 'médio'
classificar(9000)   # 'alto'
```

A ordem importa! O Python para no primeiro que der verdade.
</details>

<details>
<summary>12. for, continue e break</summary>

```python
arquivos = ["a.docx", "b.pdf", "c.docx", "PARAR.docx", "d.docx"]
for arquivo in arquivos:
    if not arquivo.endswith(".docx"):
        continue
    if arquivo == "PARAR.docx":
        break
    print(arquivo)
# a.docx
# c.docx
```
</details>

<details>
<summary>13. while</summary>

```python
tentativa = 1
while tentativa <= 3:
    print(f"Tentativa {tentativa}")
    tentativa += 1
```
</details>

<details>
<summary>14. Valor em reais para número</summary>

```python
valor = float("1.500,00".replace(".", "").replace(",", "."))
print(valor + 250)   # 1750.0
```

Primeiro tira o ponto do milhar. Depois troca a vírgula por ponto.
</details>

<details>
<summary>15. Desafio</summary>

```python
import re
import docx

documento = docx.Document("teste.docx")
texto = "\n".join(p.text for p in documento.paragraphs)

def buscar(padrao):
    achado = re.search(padrao, texto, re.IGNORECASE | re.MULTILINE)
    return achado.group(1).strip() if achado else ""

dados = {
    "segurado": buscar(r"^\s*SEGURADO\s*:\s*([^\r\n]+)"),
    "placa": buscar(r"^\s*PLACA\s*:\s*([^\r\n]+)"),
    "data_ocorrencia": buscar(r"^\s*DATA\s+DA\s+OCORR.NCIA\s*:\s*(\d{1,2}/\d{1,2}/\d{4})"),
}
print(dados)
```
</details>

---

## Sites que me ajudam

- [regex101.com](https://regex101.com): testo o regex antes de colocar no código. Escolha **Python** no menu da esquerda.
- [Documentação do `re`, em português](https://docs.python.org/pt-br/3/library/re.html)
- [Documentação do Selenium](https://selenium-python.readthedocs.io/)
- [Documentação do python-docx](https://python-docx.readthedocs.io/)
