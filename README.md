# 🛰️ JobRadar

**Busca vagas de emprego em várias plataformas, filtra pelo que interessa pra você, e (se quiser) sugere uma mensagem personalizada pra puxar assunto com o recrutador.**

Feito pra quem está em recolocação de carreira e cansou de repetir a mesma busca manual toda semana em vários sites diferentes.

> ⚠️ **O que este projeto NÃO faz:** ele não se candidata, não preenche formulário, não envia mensagem automaticamente em lugar nenhum. Ele só busca e organiza. Enviar ou não o que ele gerar é sempre você quem decide, manualmente.

---

## 📋 Índice

- [O que ele faz](#-o-que-ele-faz)
- [Pré-requisitos](#-pré-requisitos)
- [Instalação](#-instalação)
- [Como usar](#-como-usar)
- [Usando IA local para gerar mensagens (opcional)](#-usando-ia-local-para-gerar-mensagens-opcional)
- [A planilha de saída](#-a-planilha-de-saída)
- [Configuração salva](#-configuração-salva)
- [Solução de problemas](#-solução-de-problemas)
- [Perguntas frequentes](#-perguntas-frequentes)
- [Contribuindo](#-contribuindo)

---

## ✨ O que ele faz

| Etapa | Descrição |
|---|---|
| 🔎 **Busca** | Consulta LinkedIn, Indeed e Google Jobs de uma vez, via [python-jobspy](https://github.com/speedyapply/JobSpy) |
| 🎯 **Filtra** | Você define cargo, localização, e se quer só vagas remotas |
| 🤖 **Personaliza** *(opcional)* | Gera uma sugestão de mensagem por vaga com um modelo de IA rodando **localmente** (Ollama) — nada sai do seu computador |
| 📊 **Exporta** | Salva tudo num `.xlsx`, sem duplicar vagas que você já viu em execuções anteriores |

---

## 🧰 Pré-requisitos

- **Python 3.9+**
- Os pacotes listados em [`requirements.txt`](requirements.txt): `python-jobspy` e `openpyxl`
- *(opcional)* [Ollama](https://ollama.com) instalado e rodando localmente, se você quiser gerar mensagens personalizadas com IA

---

## 🚀 Instalação

```bash
# 1. Clone ou baixe este repositório
git clone <url-do-seu-fork-ou-repo>
cd jobradar

# 2. Instale as dependências
pip install -r requirements.txt --break-system-packages
```

> 💡 O `--break-system-packages` é necessário em distros Linux mais recentes (Debian/Ubuntu/Fedora) que bloqueiam `pip install` fora de um ambiente virtual. Se preferir isolar o ambiente, use um `venv`:
> ```bash
> python3 -m venv .venv
> source .venv/bin/activate
> pip install -r requirements.txt
> ```

---

## ▶️ Como usar

Rode o script e responda as perguntas — todas têm um valor padrão, então basta apertar **Enter** se quiser aceitar a sugestão:

```bash
python3 jobradar.py
```

```
? Cargo/termo de busca (padrão: Analista de Suporte):
? Localização (padrão: São Paulo, Brasil):
? Somente vagas remotas? (y/n, padrão: branco = qualquer uma):
? Quantas vagas buscar (aprox.) (padrão: 100):

? Gerar mensagem personalizada com IA local (Ollama)? (y/n, padrão: n):
```

Depois disso o script busca as vagas, mostra o progresso no terminal e salva tudo em `vagas_personalizadas.xlsx` na mesma pasta.

Rodando de novo mais tarde, ele **não duplica** vagas que já estão na planilha — só adiciona as novas.

---

## 🤖 Usando IA local para gerar mensagens (opcional)

Essa parte é totalmente opcional. Se você só quer a lista de vagas, pode pular — responda `n` na pergunta sobre IA e ignore esta seção.

Se quiser que o script sugira uma mensagem personalizada pra cada vaga (pra você revisar e mandar manualmente pro recrutador), siga os passos abaixo. Tudo roda **no seu computador**, sem mandar seus dados pra nenhum serviço externo.

### Passo 1 — Instale o Ollama

Baixe em **[ollama.com](https://ollama.com)** e siga as instruções pro seu sistema operacional (Linux, Mac ou Windows).

### Passo 2 — Baixe um modelo

Depois de instalado, num terminal separado, baixe um modelo pequeno e rápido (recomendado para rodar em notebooks comuns):

```bash
ollama pull llama3.2
```

> Se seu computador tiver mais RAM/GPU disponível, modelos maiores (`llama3.1`, `mistral`, `qwen2.5`, etc.) tendem a escrever mensagens melhores. Modelos menores são mais rápidos mas às vezes fogem um pouco das instruções.

### Passo 3 — Deixe o Ollama rodando

Na maioria das instalações, o Ollama já fica rodando em segundo plano como um serviço depois de instalado. Se não estiver, inicie manualmente:

```bash
ollama serve
```

Por padrão ele fica disponível em `http://localhost:11434` — esse já é o valor sugerido pelo JobRadar, então normalmente você não precisa mudar nada.

### Passo 4 — Rode o JobRadar e responda "sim" pra IA

```bash
python3 jobradar.py
```

Quando perguntar `Gerar mensagem personalizada com IA local (Ollama)?`, responda `s`. Em seguida:

1. Confirme (ou ajuste) a URL do Ollama e o nome do modelo baixado (ex: `llama3.2:latest`)
2. Cole um resumo curto do seu perfil profissional — quanto mais específico (cargo-alvo, principais skills, anos de experiência), melhor a mensagem gerada

O script testa a conexão com o Ollama antes de começar. Se ele não conseguir se conectar, avisa e segue a busca normalmente — só que sem mensagem personalizada (a coluna "Mensagem" fica em branco).

> ⚠️ **Sempre revise a mensagem gerada antes de enviar.** É uma sugestão de rascunho, não um texto pronto — o modelo pode errar detalhes, ser repetitivo, ou simplesmente não soar como você.

---

## 📊 A planilha de saída

O arquivo `vagas_personalizadas.xlsx` tem estas colunas:

| Coluna | Conteúdo |
|---|---|
| Vaga | Título da vaga |
| Empresa | Nome da empresa |
| Local | Localização informada na vaga |
| Remoto | Se a vaga é remota (quando essa informação está disponível) |
| Link Vaga | Link direto pra vaga original |
| Mensagem | Sugestão de mensagem gerada por IA (vazio se você não usou essa opção) |

Se você rodar o script várias vezes, ele reaproveita essa planilha e só acrescenta vagas novas (comparando pelo link). Se o arquivo estiver corrompido ou for de uma versão antiga do script com colunas diferentes, o JobRadar guarda um backup automático (`vagas_personalizadas_backup_<timestamp>.xlsx`) e começa uma planilha nova, em vez de travar.

---

## 💾 Configuração salva

Depois da primeira execução, suas respostas (cargo, localização, preferências de IA, etc.) ficam salvas em `jobradar_config.json`, na mesma pasta. Da próxima vez, o script pergunta se você quer reaproveitar essa configuração — assim você não precisa redigitar tudo toda vez.

Esse arquivo **não é versionado** (está no `.gitignore`) porque pode conter seu perfil profissional e preferências pessoais.

---

## 🔧 Solução de problemas

<details>
<summary><strong>"Faltam dependências" ao rodar o script</strong></summary>

Rode:
```bash
pip install -r requirements.txt --break-system-packages
```
</details>

<details>
<summary><strong>"Não consegui falar com o Ollama"</strong></summary>

Verifique se o Ollama está instalado e rodando (`ollama serve`) e se a URL configurada bate com onde ele está escutando (padrão: `http://localhost:11434`). Sem isso, a busca de vagas continua normalmente — só a coluna de mensagem fica em branco.
</details>

<details>
<summary><strong>"Modelo não encontrado no Ollama"</strong></summary>

O nome do modelo configurado não foi baixado ainda. Rode `ollama pull <nome-do-modelo>` (ex: `ollama pull llama3.2`) e tente de novo.
</details>

<details>
<summary><strong>"Falha ao buscar vagas"</strong></summary>

Geralmente é problema de conexão com a internet, ou um dos sites (LinkedIn/Indeed/Google) bloqueou buscas automatizadas temporariamente por excesso de requisições em pouco tempo. Espere um pouco e tente de novo, ou reduza a quantidade de vagas buscadas.
</details>

<details>
<summary><strong>"Não consegui salvar em vagas_personalizadas.xlsx"</strong></summary>

Normalmente acontece se o arquivo estiver aberto no Excel/LibreOffice/Google Sheets. Feche o programa que está com ele aberto. O script já salva automaticamente num nome alternativo pra não perder o que foi buscado nessa execução.
</details>

---

## ❓ Perguntas frequentes

**O script se candidata às vagas automaticamente?**
Não, e essa é uma decisão intencional do projeto. Ele só busca e organiza informação. Enviar candidatura ou mensagem é sempre uma ação manual sua.

**Meus dados (perfil, currículo) são enviados pra algum servidor externo?**
Não. A geração de mensagem usa o Ollama rodando localmente no seu computador — nada é enviado pra fora, exceto as buscas de vaga em si (que são requisições públicas aos sites de emprego, sem dados pessoais).

**Preciso ter uma GPU pra usar a parte de IA?**
Não é obrigatório. Modelos pequenos como `llama3.2` rodam em CPU, só um pouco mais devagar. Ter GPU ajuda bastante na velocidade, mas não é pré-requisito.

**Posso usar sem a parte de IA?**
Sim, é totalmente opcional. Responda `n` na pergunta correspondente e o script funciona só como buscador de vagas.

---

## 🤝 Contribuindo

Projeto pessoal, gratuito, feito pra ajudar quem está em busca de recolocação. Sugestões, correções e melhorias são bem-vindas via issue ou pull request.

## 📄 Licença

Veja [LICENSE](LICENSE).
