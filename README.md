# Radar de Vagas

Robô de descoberta e ranqueamento de vagas de emprego. Busca vagas via
[JobSpy](https://github.com/cullenwatson/JobSpy) nos sites listados em
`sites_busca` do `perfil.yaml` (padrão: LinkedIn e Indeed; Glassdoor e
ZipRecruiter são opcionais) e também projetos freelance no **99freelas**
(scraping com requests/BeautifulSoup, categoria em `categoria_99freelas`;
`null` desativa). Avalia a afinidade de cada vaga com o perfil usando scoring
por keywords + heurísticas (skills, senioridade, modelo de trabalho, termos de
exclusão e pequenos bônus).

Execução **manual**: em 09/10/2026 não há agendamento (crontab, `hermes cron`
ou serviço systemd) para este projeto na VM.

## Uso
```bash
pip install -r requirements.txt   # python-jobspy, pyyaml, openpyxl, beautifulsoup4, requests
cp perfil.template.yaml perfil.yaml   # preencha com seus dados (não versionado)
python radar_unified.py                    # busca completa
python radar_unified.py --termo "AI Eng"   # busca só esse termo (substitui termos_busca do perfil)
python radar_unified.py --apenas-remoto    # força modelo_trabalho=remoto (JobSpy is_remote + scoring)
python radar_unified.py --top 15           # mostra top 15 (padrão 20)
python radar_unified.py --sem-excel        # não gera a planilha
python radar_unified.py --telegram         # grava ultimo_resumo_telegram.txt (não envia)
```
Na VM já existe `.venv/` na pasta (`.venv/bin/python radar_unified.py`);
`.venv_win/` é um venv do Windows, inútil aqui.

Saída: relatório no terminal + planilha `radar_vagas_AAAAMMDD_HHMM.xlsx` na
**raiz da pasta** (não em `resultados/`), gerada localmente e não versionada —
ver `.gitignore`. Sem openpyxl, cai para CSV.

## Configurar seu perfil
Veja `CONFIGURAR_PERFIL.md` — inclui um prompt pronto para gerar o
`perfil.yaml` automaticamente a partir do seu currículo com qualquer LLM
(ChatGPT, Claude, Gemini).

## Arquivos
- `radar_unified.py` — motor principal (scraping + scoring + relatório/Excel).
- `perfil.template.yaml` — template do perfil (versionado).
- `perfil.yaml` — seu perfil real preenchido (gerado localmente, **não
  versionado**, contém dados pessoais).
- `CONFIGURAR_PERFIL.md` — guia de configuração do perfil.
- `requirements.txt` — dependências.
- Locais/não versionados: `perfil.json` (formato antigo; o código atual não o
  lê), `historico_vagas.txt` (o código tem funções de histórico, mas o `main`
  atual não as chama), `ultimo_resumo_telegram.txt`, `radar_vagas_*.xlsx`,
  `resultados/` (só planilhas antigas de 04 e 10/09 — o código atual não grava nele).
