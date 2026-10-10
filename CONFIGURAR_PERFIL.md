# 🎯 Guia de Configuração do Perfil — Radar de Vagas

Este guia explica como configurar o arquivo `perfil.yaml` para que o robô do Radar de Vagas encontre e classifique as melhores vagas para o seu perfil profissional.

---

## 🚀 Método Rápido: Gerar com Inteligência Artificial (ChatGPT, Claude ou Gemini)

Se você já tem um currículo pronto (PDF ou texto) ou quer que uma IA monte o `perfil.yaml` para você, basta copiar o prompt abaixo e colar na sua IA de preferência junto com o seu currículo.

### 📋 Prompt para a IA (Copie e Cole)

```text
Você é um especialista em recrutamento técnico e engenharia de dados.
Sua tarefa é analisar o meu currículo anexado (ou histórico profissional) e gerar o arquivo de configuração `perfil.yaml` para o meu robô rastreador de vagas (JobSpy + LinkedIn + Indeed).

O arquivo gerado DEVE seguir rigorosamente a sintaxe YAML e o esquema abaixo.

Regras de preenchimento:
1. nome: Meu nome completo extraído do currículo.
2. habilidades:
   - linguagens: apenas linguagens de programação/consulta em que tenho experiência real.
   - frameworks: bibliotecas, frameworks e plataformas (ex: React, Django, Power BI, FastAPI, Pandas).
   - ferramentas: ferramentas operacionais (ex: Git, Docker, Excel Avançado, Jira, Linux).
   - areas: áreas de atuação e domínios de negócio evidenciados no currículo (ex: Análise de Dados, BI, Logística, Finanças).
3. skills_basico: tecnologias que mencionei apenas de passagem ou nível básico.
4. senioridade: deduza o meu nível realista de mercado (ex: Júnior, Pleno, Sênior) baseado nos meus anos de experiência e projetos.
5. modelo_trabalho: [Remoto, Híbrido, Presencial] ou apenas o que for mais relevante.
6. localizacao: Cidade, UF, Brasil onde resido.
7. pais_indeed: "brazil" (ou outro país caso foco internacional).
8. paises_interesse: lista com países alvo.
9. idiomas: idiomas e níveis citados no currículo.
10. termos_busca: crie entre 6 e 10 títulos de cargos reais e frequentes no mercado brasileiro que correspondam exatamente ao meu perfil para busca no LinkedIn/Indeed (ex: "Analista de Dados Júnior", "Analista de BI").
11. sites_busca: ["linkedin", "indeed"]
12. categoria_99freelas: categoria do 99freelas mais alinhada (ex: "web-mobile-e-software") ou null se não aplicável.
13. filtros:
    - max_resultados_por_termo: 25
    - dias_recentes: 7
    - excluir_termos: lista de termos de cargos indesejados (ex: se sou Júnior/Pleno, excluir "Estágio", "Senior", "Gerente", "Head", "Diretor").

Aqui está o meu currículo / informações profissionais:
[COLE SEU CURRÍCULO OU TEXTO AQUI]

Por favor, devolva EXCLUSIVAMENTE o bloco de código YAML do perfil.yaml, pronto para salvar.
```

Depois que a IA responder:
1. Copie o YAML gerado.
2. Cole no arquivo `perfil.yaml` na raiz da pasta `radar_vagas/`.
3. Pronto! Execute o radar.

---

## 📝 Método Manual: Questionário Chave

Se preferir preencher manualmente, copie `perfil.template.yaml` para `perfil.yaml` e responda às seguintes perguntas:

### 1. Quem é você e onde você quer trabalhar?
- **Nome:** Qual o seu nome completo?
- **Localização:** Em qual cidade e estado você mora ou quer trabalhar? (ex: `"Curitiba, PR, Brasil"`)
- **Modelo:** Você aceita trabalhar Presencial, Híbrido ou apenas 100% Remoto?

### 2. Quais são as suas habilidades comprovadas?
- **Linguagens:** Quais linguagens você domina no dia a dia? (ex: Python, SQL, JavaScript)
- **Frameworks / BI:** Quais bibliotecas e ferramentas de dados você usa? (ex: Power BI, DAX, Power Query)
- **Ferramentas:** Quais softwares e utilitários você usa? (ex: Excel Avançado, Git, Docker)
- **Áreas:** Quais domínios de negócio você domina? (ex: Análise de Dados, Controladoria, Logística)

### 3. Qual o seu nível de senioridade?
- Você está buscando vagas de **Júnior**, **Pleno** ou **Sênior**?
- O robô pontua positivamente vagas alinhadas e penaliza vagas de níveis diferentes.

### 4. Quais cargos você quer que o robô busque?
- Liste de 4 a 10 variações de títulos que recrutadores usam para o seu cargo.
- *Dica:* Evite termos genéricos demais como apenas "Analista" ou "Developer". Prefira "Analista de BI Júnior", "Analista de Dados Pleno", "Python Developer".

### 5. O que você NÃO quer de jeito nenhum?
- Em `excluir_termos`, coloque palavras que devem descartar uma vaga imediatamente caso apareçam no título (ex: se não quer estágio, coloque "Estágio", "Estagiário", "Internship"; se não quer liderança, coloque "Gerente", "Head", "Lead").

---

## ⚙️ Como o Algoritmo de Afinidade Funciona

O cálculo de afinidade (0% a 100%) no `radar_unified.py` soma pontos e divide pelo máximo atingível (75 = 50 + 15 + 10), limitando o resultado entre 0 e 100:

1. **Skills (teto de 50 pontos):** cada skill de `habilidades` (todas as categorias) encontrada no título/descrição soma 5 pontos; cada item de `skills_basico` soma 2.
2. **Senioridade (até 15 pontos):** nível igual ao do perfil = 15; vaga sem nível especificado = 10; se você é Pleno, vaga Júnior = 12 e Sênior = 5; demais casos = 3.
3. **Modelo de Trabalho (até 10 pontos):** remoto aceito = 10; híbrido aceito = 8; presencial aceito = 7; demais casos = 3.
4. **Termos de exclusão (penalidade):** cada termo de `excluir_termos` no título tira 30 pontos; só na descrição, tira 15.
5. **Bônus pequenos:** open source (+3), startup/scale-up (+2), equity/stock (+2).

Não há pontuação por localização: `localizacao` só é usada como local de busca no JobSpy.

---

## 🏃 Como Executar o Radar

Após salvar o seu `perfil.yaml`:

```bash
# Rodar varredura completa e gerar planilha Excel
python radar_unified.py

# Rodar e gerar resumo formatado para o Telegram
# (só grava ultimo_resumo_telegram.txt e mostra o preview; não envia nada)
python radar_unified.py --telegram

# Buscar apenas vagas remotas (força modelo_trabalho=remoto)
python radar_unified.py --apenas-remoto

# Buscar só um termo (substitui os termos_busca do perfil)
python radar_unified.py --termo "Analista de Dados"

# Outras opções: --top N (padrão 20), --sem-excel
```

O radar roda **manualmente**; na VM não há agendamento configurado para ele (verificado em 09/10/2026).

---

## 🏁 Como Usar e Executar no Windows

### 1. Configurar o Perfil
- Copie o arquivo `perfil.template.yaml` para `perfil.yaml` (ou cole o YAML gerado pela IA a partir do prompt deste guia).
- Ajuste suas preferências de busca, skills e senioridade no `perfil.yaml`.

### 2. Passos para Executar no Windows (PowerShell / Prompt de Comando)
1. Abra o terminal na pasta do projeto no Windows.
2. Crie e ative o ambiente virtual:
   python -m venv .venv
   .venv\Scripts\activate
3. Instale as dependências:
   pip install -r requirements.txt
4. Execute o radar:
   python radar_unified.py

O pipeline é compatível com Windows. Ele coletará as vagas nos sites de `sites_busca` (padrão LinkedIn e Indeed; Glassdoor/ZipRecruiter se você incluir) e os projetos do 99freelas (se `categoria_99freelas` não for null), calculará as pontuações e gerará a planilha Excel na raiz da pasta.
