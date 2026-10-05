# Radar de Vagas

Robô de descoberta e ranqueamento de vagas de emprego. Varre LinkedIn, Indeed
e Glassdoor via [JobSpy](https://github.com/cullenwatson/JobSpy) e avalia a
afinidade de cada vaga com o perfil do candidato usando scoring por
keywords + heurísticas (linguagens, frameworks, ferramentas, senioridade,
modelo de trabalho, localização).

## Uso
```bash
pip install -r requirements.txt   # jobspy, pyyaml, openpyxl
cp perfil.template.yaml perfil.yaml   # preencha com seus dados (não versionado)
python radar_unified.py                    # busca completa
python radar_unified.py --termo "AI Eng"   # busca específica
python radar_unified.py --apenas-remoto    # só vagas remotas
python radar_unified.py --top 15           # mostra top 15
```

Saída: relatório no terminal + planilha Excel com todas as vagas ranqueadas
(gerada localmente, não versionada — ver `.gitignore`).

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
