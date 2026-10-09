# NecOrc: modelo preditivo de necessidade orçamentária

Modelo preditivo de necessidade orçamentária - Ministério das Cidades.

**Versão 1.0** (09/10/2026)

Sistema local em Python que lê a planilha atualizada (o mesmo arquivo do Excel), recalcula o modelo
preditivo com as fórmulas do Excel reescritas em Python e mostra as tabelas resumo da aba
"parametros e resultados gerais". O usuário só altera os parâmetros; as bases continuam sendo
atualizadas e tratadas no Excel, como hoje.

## 1. Instalação (uma vez)

Pré-requisitos: Python 3.10 ou superior e VS Code com a extensão Python.

No terminal do VS Code (*Terminal > New Terminal*), na pasta do projeto (Windows/PowerShell):

```
py -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
```

No VS Code: *File > Open Folder* na pasta do projeto e, se quiser, selecione o interpretador `.venv`
(*Ctrl+Shift+P > Python: Select Interpreter*).

## 2. Rotina de uso

1. Atualize a planilha no Excel como hoje (view SQL, SIAFI, NecFin CAIXA, Painel SPOA - RP3,
   acao_ajustada) e **salve** o arquivo .xlsx.
2. Copie/salve o arquivo na pasta `dados/` (configurável em `config.json`).
3. Abra o sistema:
   ```
   .\.venv\Scripts\python -m streamlit run app.py
   ```
   Para parar: Ctrl+C no terminal.
   ou, no VS Code, *Run and Debug > "Abrir sistema (Streamlit)"*. O navegador abre em
   http://localhost:8501.
4. Na barra lateral escolha a planilha (a mais recente vem primeiro). Os parâmetros começam com os
   valores salvos na aba de parâmetros; altere à vontade, os resultados recalculam na hora.
   "Restaurar parâmetros da planilha" volta aos valores do arquivo.
5. "Gerar arquivo Excel com os resultados" baixa os resumos, os parâmetros usados e o detalhe por
   contrato.

### Acesso pela rede (chefias e técnicos)

Dê dois cliques em `iniciar_sistema.bat`. A janela mostra o endereço para as outras pessoas
(`http://NOME-DO-COMPUTADOR:8501`, ou o "Network URL" com o IP). Quem estiver na mesma rede abre esse
endereço no navegador e entra com o login. O sistema fica no ar enquanto a janela estiver aberta.
Cada pessoa tem a sua própria sessão: os parâmetros que uma muda não afetam as outras.

### Usuários e senhas

O acesso é controlado pela planilha `dados/usuarios.xlsx` (não vai para o GitHub), com as colunas
`usuario | nome | perfil | senha`:

- **admin**: escolhe a planilha, recarrega e vê as opções de desenvolvedor do menu ⋮;
- **consulta**: altera os parâmetros e vê os resultados da planilha salva mais recentemente.

Mudanças na planilha de usuários valem no próximo acesso, sem reiniciar. O login continua ao apertar
F5 (código no endereço da página, `?acesso=...`); "Sair" na barra lateral encerra.

## 3. O que o sistema lê da planilha

| Aba | Intervalo | Uso |
|---|---|---|
| view_preditivo_orc_pac_2026 | colunas A..AM (pelo nome do cabeçalho) | dados dos contratos |
| SIAFI | B5:E5000 | empenhado_2026 por cod_tci |
| NecFin CAIXA | A3:L5000 | necessidade financeira (coluna M = MÁXIMO(I:J) é recalculada) |
| Painel SPOA - RP3 | C8:H50 | disponível LOA por ação |
| acao_ajustada | A1:B29 e A29:D32 | de-para de ações e ajustes pontuais |
| parametros e resultados gerais | E4:I10, L11:O11, K26:L45 | valores iniciais dos parâmetros, nomes das secretarias e lista de ações do resumo |

As colunas calculadas (M e AN..BV) **não** são lidas para o cálculo: são refeitas em Python.

## 4. Validação

```
.\.venv\Scripts\python validar.py dados\NecOrc_PAC_aprimorado.xlsx
```

Compara, com os parâmetros salvos no arquivo, todas as 36 colunas calculadas de cada contrato e as
tabelas K12:P18, K20:P20 e K26:P47 com os valores que o Excel gravou. Rode sempre que mudar uma
fórmula no Excel ou trocar a estrutura da planilha. O arquivo precisa ter sido salvo pelo Excel
depois de recalcular.

## 5. Estrutura

```
app.py              front-end (Streamlit)
assets/style.css    tema visual
.streamlit/         cores base do Streamlit
validar.py          conferência Python x Excel
config.json         pasta das planilhas
necorc/excel.py     comportamento das funções do Excel (SE, PROCV, SOMASES, ARRED, comparações)
necorc/leitura.py   leitura das abas
necorc/modelo.py    fórmulas da view (uma a uma, com a coluna do Excel comentada) e tabelas resumo
necorc/exportar.py  tabelas para tela e exportação
```

Se uma fórmula mudar no Excel, altere o trecho correspondente em `necorc/modelo.py` (cada bloco
indica a coluna, ex.: `# BQ`) e rode a validação.

## 6. Publicar só o front-end

O app roda na sua máquina. Para outras pessoas da rede acessarem:
`.\.venv\Scripts\python -m streamlit run app.py --server.address 0.0.0.0` e compartilhe `http://<seu-ip>:8501`.
Os dados continuam apenas na sua máquina.

## 7. Versão web (página publicada)

`python gerar_web.py dados\<planilha>.xlsx dados\usuarios.xlsx` gera `web/saida/necorc.html`, a página
publicada para os usuários (login com os usuários de `usuarios.xlsx`, dados cifrados com a senha).
O Python calcula o que não depende dos parâmetros; `web/motor.js` recalcula o resto na página com as
mesmas fórmulas de `necorc/modelo.py`. O botão "Gerar arquivo Excel com os resultados" gera as mesmas
4 abas da versão local (`web/vendor/` tem a biblioteca SheetJS usada para montar o .xlsx).
