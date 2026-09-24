# Padronização de Relação de Placas

Script em Python para padronizar uma planilha de relação de veículos e transportadores.

O projeto foi criado para automatizar a limpeza e padronização dos dados antes de seu uso em análises, relatórios ou processos internos.

## Objetivo

A partir de uma planilha Excel de origem, o script:

- Mantém somente as colunas necessárias:
  - `NOME DO TRANSPORTADOR`
  - `TIPO DO VEÍCULO`
  - `PLACA`
- Remove espaços extras.
- Remove quebras de linha.
- Corrige espaços duplicados.
- Padroniza nomes de transportadores por palavras-chave.
- Padroniza placas.
- Mantém transportadores que ainda não estão cadastrados, sem alterar seu nome para outro transportador.
- Gera uma nova planilha Excel com os dados tratados.

## Estrutura do projeto

```text
Padronizacao-Placas/
│
├── padronizar_relacao_placas.py
├── transportadores.py
├── README.md
├── .gitignore
│
└── Dados/
    └── planilha_fonte.xlsx
```

> As planilhas de dados não devem ser versionadas no Git. O diretório `Dados/` é destinado aos arquivos de entrada e saída locais.

## Local da fonte

Atualmente, o script procura automaticamente a planilha de origem em:

```text
C:\Users\jefferson.novaes\Downloads\Dados
```

Não é necessário informar o nome da planilha no código.

O script procura arquivos `.xlsx` nessa pasta e ignora o arquivo de saída:

```text
RELAÇÃO_DE_PLACAS_PADRONIZADA.xlsx
```

### Importante

Para evitar que a planilha errada seja processada, deve existir apenas uma planilha `.xlsx` de origem na pasta.

Se houver mais de uma, o script interrompe a execução e informa os arquivos encontrados.

## Resultado

O arquivo tratado será criado na mesma pasta da fonte:

```text
RELAÇÃO_DE_PLACAS_PADRONIZADA.xlsx
```

## Regras de padronização

O projeto utiliza palavras-chave para identificar transportadores.

Por exemplo:

```text
PACHECO
PACHECO TRANSPORTES
PACHECO & PACHECO LTDA
PACHECO E PACHECO TRANSPORTE
```

serão padronizados para:

```text
PACHECO E PACHECO
```

Da mesma forma:

```text
FAP
Fernando Augusto Pires
FERNANDO AUGUSTO PIRES LTDA
```

serão padronizados para:

```text
Fernando Augusto Pires LTDA
```

### Novas transportadoras

Quando uma nova transportadora entrar na operação, basta adicionar uma nova regra no arquivo responsável pelas regras de transportadores.

Exemplo:

```python
("NOVA LOG", "NOVA LOGISTICA LTDA"),
```

Não é necessário cadastrar todas as possíveis formas de erro de digitação.

O sistema utiliza a palavra-chave para encontrar variações do nome.

## Tratamento de espaços

O script corrige automaticamente espaços duplicados.

Exemplo:

```text
EASY  CARGO
```

vira:

```text
EASY CARGO
```

Também são removidos espaços no início e no final:

```text
  EASY CARGO
```

vira:

```text
EASY CARGO
```

Quebras de linha também são tratadas.

## Tratamento das placas

As placas são:

- convertidas para letras maiúsculas;
- têm todos os espaços removidos.

Exemplo:

```text
abc 1d23
```

vira:

```text
ABC1D23
```

## Transportadores não cadastrados

Se o script encontrar uma transportadora que ainda não possui uma regra, ele **não tenta adivinhar** qual transportadora é.

O nome original é mantido, com apenas a limpeza dos espaços.

Isso evita que uma transportadora nova seja incorretamente associada a outra.

## Requisitos

Python 3.x

Bibliotecas utilizadas:

```text
pandas
openpyxl
```

Instalação:

```powershell
pip install pandas openpyxl
```

## Execução

Abra o PowerShell ou terminal e execute:

```powershell
python padronizar_relacao_placas.py
```

O script irá:

1. Localizar a planilha de origem.
2. Ler os dados.
3. Validar as colunas necessárias.
4. Remover as colunas desnecessárias.
5. Limpar os dados.
6. Padronizar os transportadores.
7. Padronizar as placas.
8. Gerar o arquivo final.

## Versionamento com Git

O Git deve ser utilizado para versionar o código e as regras de negócio, e não as planilhas de dados.

### Quando fazer commit

Não é necessário fazer commit cada vez que o script for executado.

Um commit deve ser realizado quando houver uma alteração no projeto, por exemplo:

- inclusão de uma nova transportadora;
- alteração de uma regra de padronização;
- correção de um erro no script;
- melhoria no processo de tratamento.

### Exemplo

Ao adicionar uma nova transportadora:

```python
("NOVA LOG", "NOVA LOGISTICA LTDA"),
```

podemos fazer:

```powershell
git add .
git commit -m "Adiciona transportadora Nova Log"
git push
```

### Exemplos de mensagens de commit

```text
Cria script de padronização
Adiciona transportadora Nova Log
Corrige regra da transportadora FAP
Atualiza regras de transportadores
Corrige tratamento de placas
Melhora validação da planilha de entrada
```

## Segurança dos dados

As planilhas de origem e resultado podem conter informações operacionais.

Por isso, os arquivos Excel devem permanecer fora do versionamento do Git.

O `.gitignore` deve conter:

```gitignore
*.xlsx
*.xls
~$*
__pycache__/
*.pyc
```

Dessa forma, o Git não enviará as planilhas para o repositório.

## Fluxo recomendado

```text
Nova planilha
      │
      ▼
Pasta Dados
      │
      ▼
Executar script
      │
      ▼
Limpeza dos dados
      │
      ▼
Padronização dos transportadores
      │
      ▼
Padronização das placas
      │
      ▼
Planilha final
```

Quando houver uma nova transportadora:

```text
Nova transportadora
      │
      ▼
Adicionar regra
      │
      ▼
Testar script
      │
      ▼
Git add
      │
      ▼
Git commit
      │
      ▼
Git push
```

## Status do projeto

Projeto em desenvolvimento.

As regras de transportadores serão atualizadas conforme novas transportadoras ou novas variações de cadastro forem identificadas.
