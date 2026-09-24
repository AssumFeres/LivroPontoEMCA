LIVRO PONTO - PROFESSORES EMCA
==============================

1. OBJETIVO
-----------
O projeto lê os horários de 2026 da EMCA, relaciona a sigla da disciplina
com o(s) professor(es) da aba Base e grava os dados em SQLite.

Depois permite exportar qualquer mês para Excel.


2. REGRA PARA DISCIPLINAS COM DOIS PROFESSORES
-----------------------------------------------
Quando a coluna E (Nome Completo) da aba Base contém dois professores
separados por '/', a aula é atribuída aos dois.

Exemplo na Base:

OTB | ... | RODNEY FAGUNDES DOS SANTOS / MANUEL DE LEMOS GASPAR

Se no horário constar:

01/09/2026 | OTB | OTB

serão geradas duas linhas:

U3 | 01/09/2026 | RODNEY FAGUNDES DOS SANTOS | OTB | OTB
U3 | 01/09/2026 | MANUEL DE LEMOS GASPAR     | OTB | OTB


3. ARQUIVOS LIDOS
-----------------
AVI:
C:\Users\391665\OneDrive\EMCA\Horários\2026\01 - AVI\AVI NOTURNO 2026.xlsx
Aba: Rel FATEPI -> Turma AVI 26

CEL:
C:\Users\391665\OneDrive\EMCA\Horários\2026\02 - CEL\CEL NOTURNO 2026.xlsx
Aba: Rel FATEPI -> Turma CEL 26

GMP:
C:\Users\391665\OneDrive\EMCA\Horários\2026\03 - GMP\GMP NOTURNO 2026.xlsx
Abas: GMP01 e GMP02

BÁSICO NOTURNO:
C:\Users\391665\OneDrive\EMCA\Horários\2026\04 - BASICO\BASICO Noturno 2026.xlsx
Abas: R3, S3 e T3

BÁSICO VESPERTINO:
C:\Users\391665\OneDrive\EMCA\Horários\2026\04 - BASICO\BASICO Vespertino 2026.xlsx
Aba: U3

A aba Base é localizada mesmo que no arquivo esteja gravada como "Base "
(com espaço no final).


4. ESTRUTURA DO BANCO
---------------------
Tabela calendario_turma:
- contém os 365 dias de 2026 para cada turma.

Tabela aulas:
- turma
- turno
- data
- nome_completo
- aula_01
- aula_02
- arquivo_origem
- aba_origem

Tabela resumo_professor_mensal:
- ano
- mes
- nome_completo
- disciplinas
- turmas

Tabela log_importacao:
- registra siglas não encontradas na Base e erros de importação.


5. IMPORTANTE SOBRE AS FÓRMULAS DOS HORÁRIOS
---------------------------------------------
As colunas AULA 01 e AULA 02 são calculadas por fórmulas no Excel.
O importador usa data_only=True para ler o resultado calculado.

Por isso, antes de importar, é recomendável que os arquivos de horário
tenham sido abertos e salvos pelo Excel após a última alteração.


6. INSTALAÇÃO
-------------
Crie a pasta:

C:\Projetos\LivroPontoEMCA

Copie os arquivos deste projeto para ela.

Abra o PowerShell:

cd C:\Projetos\LivroPontoEMCA
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py


7. EXPORTAÇÃO
-------------
O Excel mensal será criado em:

C:\Projetos\LivroPontoEMCA\exportacoes

Exemplo:
LIVRO_PONTO_2026_09_SETEMBRO.xlsx

Abas:
- Livro Ponto
- Resumo Professores

Livro Ponto:
TURMA | DATA | NOME COMPLETO | AULA 01 | AULA 02

Resumo Professores:
NOME DO PROFESSOR | DISCIPLINAS DO MÊS | TURMAS DO MÊS
