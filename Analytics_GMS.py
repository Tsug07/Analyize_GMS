"""
Analytics GMS - Comparador de Relatórios de Empregados vs GMS
==============================================================
Compara o relatório geral de empregados (Empregados.pdf) com os
relatórios individuais de GMS de cada empresa para identificar
funcionários que constam no geral mas faltam no individual.
"""

import os
import re
import logging
from typing import Optional
from tkinter import filedialog, messagebox
from datetime import datetime
import customtkinter as ctk
import PyPDF2
import requests
from dotenv import load_dotenv

# ============================================================
# CONFIGURAÇÃO
# ============================================================

# Configuração de logging
_LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(_LOG_DIR, exist_ok=True)
_LOG_PATH = os.path.join(_LOG_DIR, "analytics.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(_LOG_PATH, encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)



def enviar_webhook_discord(mensagem: str, webhook_url: Optional[str] = None) -> bool:
    """Envia uma mensagem para o Discord via webhook configurado."""
    webhook_url = webhook_url or os.getenv('DISCORD_WEBHOOK_URL')
    if not webhook_url:
        logger.warning('Webhook Discord não configurado. Ignorando notificação.')
        return False

    try:
        response = requests.post(webhook_url, json={'content': mensagem}, timeout=15)
        if response.status_code == 204:
            logger.info('Notificação enviada ao Discord com sucesso.')
            return True
        logger.warning('Falha ao enviar webhook Discord: %s - %s', response.status_code, response.text)
        return False
    except Exception as e:
        logger.error('Erro ao enviar Discord webhook: %s', e)
        return False


# Carrega .env automaticamente ao iniciar o script
load_dotenv()


# ============================================================
# INTERFACE GRÁFICA (CustomTkinter)
# ============================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class AnalyticsGUI:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Analytics GMS - Comparador de Relatórios")
        self.root.geometry("750x520")
        self.root.resizable(True, True)

        # Variáveis
        self.pdf_geral_path = ""
        self.pasta_individual_path = ""

        self.setup_ui()

    def setup_ui(self):
        # Frame principal
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Título
        ctk.CTkLabel(main_frame, text="Analytics GMS",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(15, 25))

        # --- PDF Geral ---
        ctk.CTkLabel(main_frame, text="PDF Geral de Empregados:",
                     font=ctk.CTkFont(size=13, weight="bold"), anchor="w").pack(
                     fill="x", padx=20)

        pdf_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        pdf_frame.pack(fill="x", padx=20, pady=(5, 15))

        self.pdf_entry = ctk.CTkEntry(pdf_frame, placeholder_text="Selecione o PDF geral...",
                                      state="disabled", width=480)
        self.pdf_entry.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(pdf_frame, text="Selecionar PDF", width=140,
                      command=self.selecionar_pdf_geral).pack(side="right", padx=(10, 0))

        # --- Pasta Individual ---
        ctk.CTkLabel(main_frame, text="Pasta com PDFs Individuais:",
                     font=ctk.CTkFont(size=13, weight="bold"), anchor="w").pack(
                     fill="x", padx=20)

        pasta_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        pasta_frame.pack(fill="x", padx=20, pady=(5, 10))

        # --- Competência ---
        comp_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        comp_frame.pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(comp_frame, text="Competência (MM/AAAA):",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=(0, 10))

        self.competencia_entry = ctk.CTkEntry(comp_frame, placeholder_text="ex: 05/2026", width=120)
        self.competencia_entry.pack(side="left")

        ctk.CTkLabel(comp_frame, text="  Funcionários admitidos após essa competência serão ignorados.",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(side="left")

        self.pasta_entry = ctk.CTkEntry(pasta_frame, placeholder_text="Selecione a pasta...",
                                        state="disabled", width=480)
        self.pasta_entry.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(pasta_frame, text="Selecionar Pasta", width=140,
                      command=self.selecionar_pasta_individual).pack(side="right", padx=(10, 0))

        # --- Status ---
        ctk.CTkLabel(main_frame, text="Status:",
                     font=ctk.CTkFont(size=13, weight="bold"), anchor="w").pack(
                     fill="x", padx=20)

        self.status_label = ctk.CTkLabel(main_frame, text="Pronto para iniciar análise",
                                         font=ctk.CTkFont(size=12),
                                         fg_color=("gray85", "gray25"),
                                         corner_radius=6, height=40,
                                         wraplength=600)
        self.status_label.pack(fill="x", padx=20, pady=(5, 20))

        # --- Progress Bar ---
        self.progress = ctk.CTkProgressBar(main_frame, width=400)
        self.progress.pack(padx=20, pady=(0, 15))
        self.progress.set(0)

        # --- Botões ---
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(pady=(0, 15))

        self.executar_btn = ctk.CTkButton(button_frame, text="Executar Análise",
                                          font=ctk.CTkFont(size=14, weight="bold"),
                                          height=40, width=220,
                                          command=self.executar_analise,
                                          state="disabled")
        self.executar_btn.pack(side="left", padx=(0, 10))

        ctk.CTkButton(button_frame, text="Testar Webhook",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      height=40, width=220,
                      fg_color="#5865F2",
                      hover_color="#4752C4",
                      command=self.testar_webhook).pack(side="left", padx=(0, 10))

        self.verificar_nomes_btn = ctk.CTkButton(button_frame, text="Verificar Nomes",
                                                  font=ctk.CTkFont(size=14, weight="bold"),
                                                  height=40, width=220,
                                                  fg_color="#2E8B57",
                                                  hover_color="#1F6B3E",
                                                  command=self.verificar_nomes,
                                                  state="disabled")
        self.verificar_nomes_btn.pack(side="left")

    def selecionar_pdf_geral(self):
        filename = filedialog.askopenfilename(
            title="Selecionar PDF Geral de Empregados",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.pdf_geral_path = filename
            self.pdf_entry.configure(state="normal")
            self.pdf_entry.delete(0, "end")
            self.pdf_entry.insert(0, filename)
            self.pdf_entry.configure(state="disabled")
            self.verificar_campos()

    def selecionar_pasta_individual(self):
        dirname = filedialog.askdirectory(
            title="Selecionar Pasta com PDFs Individuais"
        )
        if dirname:
            self.pasta_individual_path = dirname
            self.pasta_entry.configure(state="normal")
            self.pasta_entry.delete(0, "end")
            self.pasta_entry.insert(0, dirname)
            self.pasta_entry.configure(state="disabled")
            self.verificar_campos()

    def verificar_campos(self):
        if self.pdf_geral_path and self.pasta_individual_path:
            self.executar_btn.configure(state="normal")
        else:
            self.executar_btn.configure(state="disabled")
        if self.pasta_individual_path:
            self.verificar_nomes_btn.configure(state="normal")
        else:
            self.verificar_nomes_btn.configure(state="disabled")

    def executar_analise(self):
        if not os.path.exists(self.pdf_geral_path):
            messagebox.showerror("Erro", "O arquivo PDF geral selecionado não existe!")
            return

        if not os.path.exists(self.pasta_individual_path):
            messagebox.showerror("Erro", "A pasta selecionada não existe!")
            return

        # Validar competência se preenchida
        competencia = self.competencia_entry.get().strip()
        if competencia:
            import re as _re
            if not _re.match(r'^\d{2}/\d{4}$', competencia):
                messagebox.showerror("Erro", "Competência inválida. Use o formato MM/AAAA (ex: 05/2026).")
                return

        try:
            self.status_label.configure(text="Iniciando análise...")
            self.progress.set(0)
            self.root.update()

            resultado = executar_analise_completa(
                self.pdf_geral_path,
                self.pasta_individual_path,
                self.atualizar_progresso,
                competencia=competencia or None
            )

            self.progress.set(1)
            self.status_label.configure(text=f"Análise concluída!")

            messagebox.showinfo("Sucesso", f"Análise concluída!\n\n{resultado}")

        except Exception as e:
            self.status_label.configure(text=f"Erro: {str(e)}")
            messagebox.showerror("Erro", f"Ocorreu um erro durante a análise:\n\n{str(e)}")

    def atualizar_progresso(self, valor, mensagem):
        self.progress.set(valor / 100)
        self.status_label.configure(text=mensagem)
        self.root.update()

    def testar_webhook(self):
        webhook_url = os.getenv('DISCORD_WEBHOOK_URL')
        if not webhook_url:
            messagebox.showwarning(
                "Webhook não configurado",
                "Nenhuma URL de Discord webhook encontrada em DISCORD_WEBHOOK_URL.\n"
                "Verifique seu arquivo .env."
            )
            return

        self.status_label.configure(text="Enviando mensagem de teste para o Discord...")
        self.root.update()

        sucesso = enviar_webhook_discord(
            "🚀 Teste de webhook do Analytics GMS: conexão bem-sucedida.",
            webhook_url=webhook_url
        )

        if sucesso:
            messagebox.showinfo(
                "Webhook enviado",
                "A mensagem de teste foi enviada com sucesso para o Discord."
            )
            self.status_label.configure(text="Teste de webhook concluído com sucesso.")
        else:
            messagebox.showerror(
                "Falha no webhook",
                "Não foi possível enviar a mensagem de teste. Verifique o URL do webhook e a conexão.")
            self.status_label.configure(text="Falha ao enviar mensagem de teste para o Discord.")

        self.root.update()

    def verificar_nomes(self):
        if not os.path.exists(self.pasta_individual_path):
            messagebox.showerror("Erro", "A pasta selecionada não existe!")
            return

        self.status_label.configure(text="Verificando nomes dos arquivos vs conteúdo...")
        self.progress.set(0)
        self.root.update()

        divergencias = verificar_nomes_arquivos_vs_conteudo(
            self.pasta_individual_path,
            self.atualizar_progresso
        )

        self.progress.set(1)

        if not divergencias:
            self.status_label.configure(text="Verificação concluída: todos os nomes conferem!")
            enviar_webhook_discord("✅ Verificação de nomes GMS concluída: todos os arquivos conferem com o conteúdo interno.")
            messagebox.showinfo("Verificação de Nomes", "Todos os arquivos conferem com o conteúdo interno.")
            return

        # Montar relatório de divergências
        linhas = [f"Encontradas {len(divergencias)} divergência(s):\n"]
        for d in divergencias:
            linhas.append(f"Arquivo:   {d['arquivo']}")
            linhas.append(f"Nome arquivo: {d['nome_arquivo']}")
            linhas.append(f"Nome interno: {d['nome_interno']}")
            linhas.append("")

        texto = "\n".join(linhas)

        # Salvar em arquivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        script_dir = os.path.dirname(os.path.abspath(__file__))
        pasta_faltantes = os.path.join(script_dir, "faltantes")
        os.makedirs(pasta_faltantes, exist_ok=True)
        output_path = os.path.join(pasta_faltantes, f"divergencias_nomes_{timestamp}.txt")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(texto)

        # Notificar Discord com detalhes das divergências
        linhas_discord = [f"⚠️ <@1285316758009544844> Verificação de nomes GMS: {len(divergencias)} divergência(s) encontrada(s).\n"]
        for d in divergencias:
            linhas_discord.append(f"📄 **Arquivo:** {d['arquivo']}")
            linhas_discord.append(f"   Nome arquivo: `{d['nome_arquivo']}`")
            linhas_discord.append(f"   Nome interno: `{d['nome_interno']}`")
        mensagem_discord = "\n".join(linhas_discord)
        # Discord limita mensagens a 2000 caracteres
        if len(mensagem_discord) > 1900:
            mensagem_discord = mensagem_discord[:1900] + f"\n... e mais divergências. Veja o relatório completo."
        enviar_webhook_discord(mensagem_discord)

        self.status_label.configure(text=f"Verificação concluída: {len(divergencias)} divergência(s) encontrada(s).")
        messagebox.showwarning(
            "Divergências Encontradas",
            f"{len(divergencias)} arquivo(s) com nome diferente do conteúdo interno.\n\n"
            f"Relatório salvo em:\n{os.path.basename(output_path)}\n\n"
            f"{texto[:800]}{'...' if len(texto) > 800 else ''}"
        )

    def run(self):
        self.root.mainloop()


# ============================================================
# UTILIDADES DE NORMALIZAÇÃO
# ============================================================

def normalizar_nome(nome: str) -> str:
    """
    Normaliza nome para comparação:
    - Remove espaços extras
    - Converte para maiúsculas
    - Remove acentos comuns que podem diferir entre PDFs
    """
    nome = nome.strip().upper()
    nome = re.sub(r'\s+', ' ', nome)
    return nome


# ============================================================
# ETAPA 1A: LEITURA DO PDF DE EMPREGADOS (RELATÓRIO GERAL)
# ============================================================

def extrair_texto_pdf(caminho_pdf: str) -> list[str]:
    """Extrai texto de cada página do PDF. Retorna lista de textos por página."""
    paginas = []
    with open(caminho_pdf, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            texto = page.extract_text()
            if texto:
                paginas.append(texto)
    return paginas


def _extrair_nome_do_trecho(trecho: str) -> str:
    """
    Dado um trecho como "ELENI TEREZA RODRIGUES OPERADORA DE CAIXA",
    separa o nome do funcionário do cargo.

    Lógica: percorre as palavras de trás para frente. As últimas palavras
    formam o cargo. Quando encontramos uma sequência que não parece cargo
    (nomes próprios), paramos — o restante é o nome.

    Palavras comuns em cargos (não em nomes de pessoas):
    """
    # Palavras que indicam que ainda estamos no cargo (lidas de trás pra frente)
    palavras_cargo = {
        'ADMINISTRATIVO', 'ADMINISTRATIVA', 'ADMINISTRADOR', 'ADMISTRATIVO', 'ADMINSITRATIVO',
        'AJUDANTE', 'ALMOXARIFADO', 'ANALISTA', 'ASSISTENTE', 'ATENDENTE',
        'AUXILIAR', 'AUX', 'BALCÃO', 'BALCONISTA', 'CAIXA', 'CHEFE',
        'COMERCIAL', 'CONFERENTE', 'CONSULTOR', 'CONSULTORA', 'COORDENADOR', 'COORDENADORA',
        'DEPARTAMENTO', 'DESIGNER', 'DIGITAL', 'DIRETOR', 'DIRETORA',
        'ELETRICISTA', 'ELETRONICA', 'EMBALADORA', 'ENCARREGADO', 'ENFERMEIRO',
        'ESCRITORIO', 'ESCRITÓRIO', 'ESPECIALISTA', 'ESTAMPARIA', 'ESTOQUE', 'ESTOQUISTA',
        'EXPEDICAO', 'EXTERNO', 'FARMACIA', 'FARMACEUTICO', 'FINANCEIRO', 'FISCAL',
        'FRANQUIA', 'GERAL', 'GERAIS', 'GERIAS', 'GERENTE', 'GESTÃO',
        'IMPRESSOR', 'INFOR', 'INSTALACAO', 'INTERIOR', 'INTERNO',
        'LICITAÇÃO', 'LIMPEZA', 'LINHA', 'LOGISTICA', 'LOJA',
        'MAQUINAS', 'MECANICO', 'MARKETING', 'MONTADOR', 'MONITORAMENTO', 'MOTORISTA',
        'OPERACIONAL', 'OPERADOR', 'OPERADORA', 'PESSOAL', 'PLANEJAMENTO',
        'PLENO', 'PRODUCAO', 'PRODUÇÃO', 'PROGRAMADOR', 'PROMOTOR',
        'RECEPCIONISTA', 'REFRIGERACAO', 'REFRIGERAÇÃO', 'REPOSITOR', 'RESPONSAVEL',
        'SECRETARIA', 'SECCIONADORA', 'SENIOR', 'SERVIÇOS', 'SERVICOS',
        'SUPERVISOR', 'SUPERVISORA', 'SUPORTE', 'SUPRIMENTOS',
        'TECNICO', 'TÉCNICO', 'VENDAS', 'VENDEDOR', 'VENDEDORA',
        'ACOUGUEIRO', 'BICICLET', 'CFTV',
        'LAVADOR', 'VEICULOS', 'VEICULO', 'DESCARGA', 'CARGA', 'COSTUREIRO', 'COSTUREIRA',
        'PERFUMISTA', 'PADEIRO', 'PADEIRA', 'GARCOM', 'GARÇOM', 'COPEIRA', 'COPEIRO',
        'PORTEIRO', 'VIGIA', 'ZELADOR', 'ZELADORA', 'FAXINEIRO', 'FAXINEIRA',
        'INSTRUMENTADOR', 'INSTRUMENTADORA', 'TORNEIRO', 'SOLDADOR', 'PINTOR', 'PINTORA',
        'SUBGERENTE', 'SUB', 'CORTADOR', 'CORTADORA', 'PEIXEIRO', 'PEIXEIRA', 'COMERCIO',
        'TEC', 'TECNICO', 'SEGURANCA', 'SEGURANÇA', 'SEG', 'INFORMATICA', 'INFORMÁTICA',
        'AUTOMACAO', 'AUTOMAÇÃO', 'AUTOMOCAO', 'SISTEMAS', 'NIVEL', 'NÍVEL', 'TRABALHO', 'A',
        'ASSIST', 'ASSISTENTE', 'ADMNISTRATIVO', 'OP.DE', 'TECIDO', 'ROUPAS',
        'SERV', 'GERAL',
        'OP', 'I', 'II', 'III', 'IV', 'IX',
        'JUNIOR', 'JR', 'PL', 'SR', 'ADM', 'COM', 'PÓS', 'TI', 'RH',
        '(A)', '(FARMACIA)',
    }

    palavras = trecho.split()
    if not palavras:
        return ""

    # Percorre de trás para frente marcando todas as palavras de cargo,
    # mesmo que haja palavras de nome intercaladas no meio (ex: "MOREIRA PERFUMISTA- FARMACIA").
    # Estratégia: encontra o índice mais à esquerda onde ainda há uma sequência
    # contígua de cargos a partir do final.
    eh_cargo = []
    for palavra in palavras:
        p = palavra.upper().strip('-.,() ')
        # Remove sufixo "(A)" ou "(a)" colado: "OPERADOR(A)" -> "OPERADOR", "VENDEDOR(A)" -> "VENDEDOR"
        p_sem_gen = re.sub(r'\([Aa]\)$', '', p).strip()
        # Abreviações com ponto: "TEC.", "AUX.", "OP.DE" -> pega parte antes do ponto
        p_base = p.split('.')[0]
        eh_cargo.append(
            p in palavras_cargo or
            p_sem_gen in palavras_cargo or
            p_base in palavras_cargo or
            p.startswith('AUX') or
            p_base.startswith('AUX') or
            p_base.startswith('TEC') or
            p_sem_gen.startswith('VENDEDOR') or
            p_sem_gen.startswith('OPERADOR') or
            p == '(A)' or p == '(FARMACIA)'
        )

    # JUNIOR/JR só é cargo quando há outro cargo verdadeiro à sua direita imediata.
    # Se JUNIOR vem antes do cargo principal (ex: "MAIA JUNIOR VENDEDOR"), ele é nome.
    # Se JUNIOR vem depois do cargo (ex: "PROGRAMADOR JUNIOR"), ele é qualificador de cargo.
    # JUNIOR/JR é cargo quando há outro cargo verdadeiro IMEDIATAMENTE à sua esquerda
    # (ex: "PROGRAMADOR JUNIOR" — PROGRAMADOR é cargo, então JUNIOR qualifica o cargo).
    # Se não há cargo à esquerda (ex: "MAIA JUNIOR VENDEDOR"), JUNIOR é parte do nome.
    sufixos_nivel = {'JUNIOR', 'JR'}
    for i, palavra in enumerate(palavras):
        p = palavra.upper().strip('-.,() ')
        if p in sufixos_nivel and eh_cargo[i]:
            cargo_real_a_esquerda = any(
                eh_cargo[j] and palavras[j].upper().strip('-.,() ') not in sufixos_nivel
                for j in range(0, i)
            )
            if not cargo_real_a_esquerda:
                # Não há cargo antes do JUNIOR — é parte do nome
                eh_cargo[i] = False

    # Preposições que conectam palavras de cargo mas também aparecem em nomes.
    # Marcamos como "cargo condicional" — só são tratadas como cargo se estiverem
    # dentro de um bloco de cargos (i.e., há cargo à direita).
    preposicoes = {'DE', 'DO', 'DA', 'DOS', 'DAS', 'E', 'EM', 'NO', 'NA'}

    # Marca preposições que estão ENTRE cargos no final como cargo também.
    # Estratégia: percorre da direita, rastreia se há cargo à direita e se há cargo
    # à esquerda (lookahead). Se sim, a preposição é parte do cargo.
    tem_cargo_a_direita = False
    for i in range(len(palavras) - 1, -1, -1):
        p = palavras[i].upper().strip('-.,() ')
        if eh_cargo[i]:
            tem_cargo_a_direita = True
        elif p in preposicoes and tem_cargo_a_direita:
            cargo_verdadeiro_esquerda = any(eh_cargo[j] for j in range(0, i))
            if cargo_verdadeiro_esquerda:
                eh_cargo[i] = True
                tem_cargo_a_direita = True
            else:
                tem_cargo_a_direita = False
        else:
            tem_cargo_a_direita = False

    # Encontra o maior bloco contíguo de cargos no final
    idx_corte = len(palavras)
    for i in range(len(palavras) - 1, -1, -1):
        p_upper = palavras[i].upper().strip('-.,() ')
        if eh_cargo[i]:
            idx_corte = i
        elif p_upper in preposicoes and idx_corte < len(palavras):
            # Preposição entre cargos: trata como transparente (não quebra o bloco)
            idx_corte = i
        else:
            # Se a palavra não é cargo mas há palavras de cargo depois dela,
            # verifica se é apenas uma palavra de nome isolada entre cargos
            # (ex: MOREIRA PERFUMISTA- FARMACIA — MOREIRA é nome, mas está antes de cargo)
            cargos_depois = sum(1 for j in range(i + 1, len(palavras)) if eh_cargo[j])
            if cargos_depois >= 1 and i < idx_corte - 1:
                continue
            break

    # O nome são as palavras antes do cargo
    nome = ' '.join(palavras[:idx_corte])
    # Remove traços e pontuação solta no final
    nome = re.sub(r'[\-\.,;]+$', '', nome).strip()
    return nome


def extrair_funcionarios_empregados(caminho_pdf: str, competencia: str = None) -> dict:
    """
    competencia: string no formato 'MM/AAAA'. Quando informada, funcionários
    admitidos após essa competência são ignorados (ex: '05/2026' exclui admissões
    a partir de 01/06/2026).
    """
    """
    Lê o PDF de Empregados e extrai os funcionários agrupados por empresa.

    Formato esperado por página:
        NOME DA EMPRESA                              Página: X/Y
        RELAÇÃO DE EMPREGADOS I                      Emissão:...
        ...
        CódigoNome Cargo Categoria Hor.NFNDADMISSÃO SINOPT
        5037ELENI TEREZA RODRIGUES OPERADORA DE CAIXA Mensalista 180,00 ...
        ...
        Total de empregados: N

    Retorno:
    {
        "NOME DA EMPRESA": ["FULANO DA SILVA", "CICLANO SOUZA", ...],
        ...
    }
    """
    # Converter competencia "MM/AAAA" para data limite: primeiro dia do mês seguinte
    from datetime import date
    data_limite = None
    if competencia:
        try:
            mes, ano = int(competencia[:2]), int(competencia[3:])
            mes_seguinte = mes + 1 if mes < 12 else 1
            ano_seguinte = ano if mes < 12 else ano + 1
            data_limite = date(ano_seguinte, mes_seguinte, 1)
        except Exception:
            logger.warning("Competência inválida: %s. Ignorando filtro.", competencia)

    paginas = extrair_texto_pdf(caminho_pdf)
    empresas = {}

    for i, texto in enumerate(paginas):
        linhas = texto.split('\n')
        if not linhas:
            continue

        # Nome da empresa é a primeira linha (antes de "Página:")
        nome_empresa = linhas[0].split("Página:")[0].strip()
        if not nome_empresa:
            logger.warning(f"Página {i+1}: não foi possível extrair nome da empresa")
            continue

        nome_empresa = normalizar_nome(nome_empresa)

        # Encontrar funcionários entre o cabeçalho e "Total de empregados"
        funcionarios = []
        em_lista = False

        for linha in linhas:
            if "CódigoNome" in linha.replace(" ", "") or ("Código" in linha and "Nome" in linha and "Cargo" in linha):
                em_lista = True
                continue
            if "Total de empregados" in linha:
                em_lista = False
                continue
            if not em_lista:
                continue

            # Formato: "5037NOME CARGO Mensalista 220,00 NN DD/MM/AAAA SN"
            match_categoria = re.match(
                r'^(\d+)(.+?)\s+(Mensalista|Horista|Diarista|Tarefeiro)\s+(\d+[,.]\d{2})\s*\d*(\d{2}/\d{2}/\d{4})',
                linha, re.IGNORECASE
            )
            if match_categoria:
                trecho = match_categoria.group(2).strip()
                data_admissao_str = match_categoria.group(5)
                nome_func = _extrair_nome_do_trecho(trecho)
                if not nome_func:
                    continue

                # Filtrar por competência: ignorar admissões após o mês da competência
                if data_limite:
                    try:
                        d, m, a = data_admissao_str.split('/')
                        data_adm = date(int(a), int(m), int(d))
                        if data_adm >= data_limite:
                            logger.debug("Ignorando %s: admissão %s após competência %s",
                                         nome_func, data_admissao_str, competencia)
                            continue
                    except Exception:
                        pass

                funcionarios.append(normalizar_nome(nome_func))

        if funcionarios:
            if nome_empresa in empresas:
                # Empresa já registrada: ignora páginas duplicadas (inativos/histórico)
                logger.debug(f"Página {i+1}: empresa '{nome_empresa}' já registrada, ignorando duplicata")
            else:
                empresas[nome_empresa] = funcionarios
            logger.debug(f"Empresa '{nome_empresa}': {len(funcionarios)} funcionário(s)")
        else:
            logger.warning(f"Página {i+1}: nenhum funcionário encontrado para '{nome_empresa}'")

    logger.info(f"Empregados.pdf: {len(empresas)} empresas extraídas")
    return empresas


# ============================================================
# ETAPA 1B: LEITURA DOS PDFs INDIVIDUAIS DE GMS
# ============================================================

def extrair_funcionarios_gms(caminho_pdf: str) -> list[str]:
    """
    Lê um relatório individual de GMS e extrai os nomes dos funcionários.

    Formato esperado:
        Nº de Ordem  Nomes dos Associados  PATRONAL  MENSAL  Razão Social
        1 CRISTIANO DA SILVA 8,00 45,00
        2 MARIA DA PENHA VENTURA 8,00 0,00
        ...

    Retorno: ["CRISTIANO DA SILVA", "MARIA DA PENHA VENTURA", ...]
    """
    paginas = extrair_texto_pdf(caminho_pdf)
    funcionarios = []

    for texto in paginas:
        linhas = texto.split('\n')

        for linha in linhas:
            # Linhas de funcionário: número de ordem + nome + valores monetários
            # Ex: "1 CRISTIANO DA SILVA 8,00 45,00"
            match = re.match(
                r'^\s*(\d+)\s+([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ][A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ\s\.]+?)\s+\d+[,\.]\d{2}',
                linha
            )
            if match:
                nome = normalizar_nome(match.group(2))
                if nome and len(nome) > 3:  # Filtrar ruído
                    funcionarios.append(nome)

    return funcionarios


# ============================================================
# VERIFICAÇÃO: NOME DO ARQUIVO vs CONTEÚDO INTERNO
# ============================================================

def _limpar_nome_pdf(nome: str) -> str:
    """Remove espaços extras que o leitor de PDF insere dentro das palavras.
    Ex: 'PONTU AL SHOPPING CHOCOLA TES L TDA' -> 'PONTUAL SHOPPING CHOCOLATES LTDA'
    Estratégia: une tokens curtos (<=3 chars) ao token anterior se juntos formam palavra coerente.
    """
    tokens = nome.split()
    resultado = []
    i = 0
    while i < len(tokens):
        token = tokens[i]
        # Une ao token anterior se ambos são curtos (fragmento de palavra quebrada pelo PDF)
        # Ex: "CHOCOLA" + "TES" -> "CHOCOLATES", "L" + "TDA" -> "LTDA"
        if resultado and token.isalpha() and len(resultado[-1]) <= 7 and len(token) <= 4:
            resultado[-1] = resultado[-1] + token
        else:
            resultado.append(token)
        i += 1
    return ' '.join(resultado)


def extrair_nome_empresa_do_pdf_gms(caminho_pdf: str) -> str:
    """
    Extrai o nome da empresa de dentro do PDF individual de GMS.
    O nome da empresa aparece no cabeçalho logo antes de 'Sindicato dos Empregados',
    repetido 3 vezes. A segunda ocorrência (após o endereço da empresa) é a mais limpa.
    """
    paginas = extrair_texto_pdf(caminho_pdf)
    if not paginas:
        return ""

    texto = paginas[0]

    # O nome da empresa aparece antes de "Sindicato dos Empregados" no cabeçalho.
    # Formato: "NOME DA EMPRESA\nSindicato dos Empregados..."
    match = re.search(r'([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ][^\n]+)\nSindicato\s+dos\s+Empregados', texto, re.IGNORECASE)
    if match:
        nome_bruto = match.group(1).strip()
        # Remove sufixo de endereço que pode vir colado (ex: "NOME | Rua X")
        nome_bruto = re.split(r'\s*[\|]\s*', nome_bruto)[0].strip()
        nome_limpo = _limpar_nome_pdf(nome_bruto)
        nome_norm = normalizar_nome(nome_limpo)
        if len(nome_norm) > 3:
            return nome_norm

    # Fallback: pega a terceira ocorrência do nome — aparece após "CARIMBO PADRONIZADO"
    match2 = re.search(r'CARIMBO\s+P\s*ADRONIZADO\s*\nCNPJ[^\n]+\n([^\n]+)', texto, re.IGNORECASE)
    if match2:
        nome_bruto = match2.group(1).strip()
        nome_limpo = _limpar_nome_pdf(nome_bruto)
        nome_norm = normalizar_nome(nome_limpo)
        if len(nome_norm) > 3:
            return nome_norm

    return ""


def verificar_nomes_arquivos_vs_conteudo(diretorio: str, callback_progresso=None) -> list:
    """
    Para cada PDF da pasta, compara o nome extraído do nome do arquivo
    com o nome da empresa encontrado dentro do conteúdo do PDF.

    Retorna lista de dicts com as divergências encontradas.
    """
    arquivos = [f for f in os.listdir(diretorio) if f.endswith(".pdf")]
    total = len(arquivos)
    divergencias = []

    for i, arquivo in enumerate(arquivos):
        if callback_progresso:
            callback_progresso(int((i / total) * 100), f"Verificando {arquivo}...")

        nome_base = arquivo.replace(".pdf", "")
        if re.match(r'^\d+-.+-\d{6}$', nome_base):
            partes = nome_base.split("-", 1)
            resto = partes[1].rsplit("-", 1)
            if len(resto) < 2:
                continue
            nome_arquivo = normalizar_nome(resto[0].strip())
        elif nome_base.startswith("Relatorio_"):
            partes = nome_base.split("_", 2)
            if len(partes) < 3:
                continue
            nome_arquivo = normalizar_nome(partes[2].replace("_", " "))
        else:
            continue
        caminho = os.path.join(diretorio, arquivo)

        try:
            nome_interno = extrair_nome_empresa_do_pdf_gms(caminho)
        except Exception as e:
            logger.warning("Erro ao ler %s: %s", arquivo, e)
            nome_interno = ""

        if not nome_interno:
            logger.warning("Não foi possível extrair nome interno de: %s", arquivo)
            continue

        # Compara removendo todos os espaços — o PDF quebra palavras com espaços extras
        # ex: "DROGARIA SENALTDA" vs "DROGARIA SENA LTDA" → ambos viram "DROGARIASSENALTDA"
        # Também ignora sufixos como "- ME", "- EPP", "LTDA", "EIRELI" que podem estar truncados
        def compactar(s):
            return re.sub(r'[\s\-\.\'\,]', '', s).upper()

        c_arquivo = compactar(nome_arquivo)
        c_interno = compactar(nome_interno)

        # Considera match se um começa com os primeiros 15 chars do outro (truncamento)
        prefixo = min(15, len(c_arquivo), len(c_interno))
        if c_arquivo[:prefixo] != c_interno[:prefixo] and c_arquivo not in c_interno and c_interno not in c_arquivo:
            divergencias.append({
                "arquivo": arquivo,
                "nome_arquivo": nome_arquivo,
                "nome_interno": nome_interno,
            })
            logger.warning("Divergência: arquivo='%s' | interno='%s'", nome_arquivo, nome_interno)

    logger.info("Verificação concluída. %d divergência(s) em %d arquivo(s).", len(divergencias), total)
    return divergencias


# ============================================================
# ETAPA 2: IDENTIFICAÇÃO DOS ARQUIVOS INDIVIDUAIS
# ============================================================

def mapear_arquivos_individuais(diretorio: str) -> dict:
    """
    Varre a pasta de relatórios e mapeia cada empresa ao seu arquivo.

    Retorno:
    {
        "CANELLA E SANTOS CONTABILIDADE EIRELI": {
            "arquivo": "Relatorio_102_CANELLA_E_SANTOS_CONTABILIDADE_EIRELI.pdf",
            "caminho": "...",
            "codigo": "102",
            "nome_empresa": "CANELLA E SANTOS CONTABILIDADE EIRELI"
        },
        ...
    }
    """
    mapa = {}
    for arquivo in os.listdir(diretorio):
        if not arquivo.endswith(".pdf"):
            continue

        nome_base = arquivo.replace(".pdf", "")

        # Formato novo: [CODIGO]-[NOME]-[MMAAAA].pdf  ex: 854-DROGARIA SENA LTDA-052026.pdf
        # Formato antigo: Relatorio_[CODIGO]_[NOME].pdf
        if re.match(r'^\d+-.+-\d{6}$', nome_base):
            partes = nome_base.split("-", 1)          # separa código do resto
            codigo = partes[0]
            resto = partes[1].rsplit("-", 1)           # separa nome do período (último traço)
            if len(resto) < 2:
                logger.warning(f"Arquivo com formato inesperado: {arquivo}")
                continue
            nome_empresa = resto[0].strip()
        elif nome_base.startswith("Relatorio_"):
            partes = nome_base.split("_", 2)
            if len(partes) < 3:
                logger.warning(f"Arquivo com formato inesperado: {arquivo}")
                continue
            codigo = partes[1]
            nome_empresa = partes[2].replace("_", " ")
        else:
            continue

        nome_empresa_norm = normalizar_nome(nome_empresa)
        mapa[nome_empresa_norm] = {
            "arquivo": arquivo,
            "caminho": os.path.join(diretorio, arquivo),
            "codigo": codigo,
            "nome_empresa": nome_empresa_norm
        }

    logger.info(f"Mapeados {len(mapa)} arquivos individuais de GMS.")
    return mapa


# ============================================================
# ETAPA 3: COMPARAÇÃO
# ============================================================

def encontrar_empresa_no_mapa(nome_empresa_geral: str, mapa_arquivos: dict) -> dict | None:
    """
    Tenta encontrar a empresa do relatório geral no mapa de arquivos.
    Usa correspondência flexível para lidar com diferenças de formatação.
    """
    nome = normalizar_nome(nome_empresa_geral)

    # Correspondência exata
    if nome in mapa_arquivos:
        return mapa_arquivos[nome]

    # Correspondência parcial (contém)
    for nome_mapa, info in mapa_arquivos.items():
        if nome in nome_mapa or nome_mapa in nome:
            logger.info(f"Correspondência parcial: '{nome_empresa_geral}' -> '{nome_mapa}'")
            return info

    return None


def comparar_funcionarios(funcionarios_geral: dict, mapa_arquivos: dict) -> list:
    """
    Compara os funcionários do relatório de empregados com os dos GMS individuais.
    A comparação é feita APENAS para empresas que têm arquivos individuais.
    A comparação é feita pelo NOME do funcionário (normalizado).

    Retorno: lista de dicts com os funcionários faltantes.
    """
    faltantes = []

    # Iterar apenas sobre empresas que têm arquivo GMS individual
    for nome_empresa_gms, info_empresa in mapa_arquivos.items():
        # Procurar a empresa nos funcionários gerais com correspondência flexível
        nome_geral_encontrado = None
        nomes_funcionarios_geral = None

        # Procura por correspondência exata
        if nome_empresa_gms in funcionarios_geral:
            nome_geral_encontrado = nome_empresa_gms
            nomes_funcionarios_geral = funcionarios_geral[nome_empresa_gms]
        else:
            # Procura por correspondência parcial
            for nome_geral, nomes in funcionarios_geral.items():
                nome_norm_gms = normalizar_nome(nome_empresa_gms)
                nome_norm_geral = normalizar_nome(nome_geral)
                if nome_norm_gms in nome_norm_geral or nome_norm_geral in nome_norm_gms:
                    nome_geral_encontrado = nome_geral
                    nomes_funcionarios_geral = nomes
                    logger.info(f"Correspondência parcial: '{info_empresa['nome_empresa']}' -> '{nome_geral}'")
                    break

        if nomes_funcionarios_geral is None:
            logger.debug(f"Empresa com relatório individual não encontrada no PDF geral: {nome_empresa_gms}")
            continue

        # Extrair nomes do GMS individual
        nomes_gms = extrair_funcionarios_gms(info_empresa["caminho"])
        set_gms = {normalizar_nome(n) for n in nomes_gms}

        # Verificar quem está faltando
        for nome_func in nomes_funcionarios_geral:
            nome_normalizado = normalizar_nome(nome_func)
            if nome_normalizado not in set_gms:
                faltantes.append({
                    "codigo_empresa": info_empresa["codigo"],
                    "nome_empresa": info_empresa["nome_empresa"],
                    "funcionario_nome": nome_func,
                    "motivo": "Funcionário não encontrado no GMS individual"
                })

    logger.info(f"Comparação finalizada. {len(faltantes)} funcionário(s) faltante(s).")
    return faltantes


# ============================================================
# ETAPA 4: GERAÇÃO DO RELATÓRIO DE FALTANTES
# ============================================================

def gerar_relatorio_faltantes(faltantes: list, caminho_saida: str):
    """Gera arquivo TXT com os funcionários faltantes, organizado por empresa."""
    # Agrupar por empresa
    por_empresa = {}
    for f in faltantes:
        chave = (f["codigo_empresa"], f["nome_empresa"])
        if chave not in por_empresa:
            por_empresa[chave] = []
        por_empresa[chave].append(f)

    with open(caminho_saida, "w", encoding="utf-8") as arq:
        arq.write("=" * 80 + "\n")
        arq.write("RELATÓRIO DE FUNCIONÁRIOS FALTANTES NA GMS\n")
        arq.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        arq.write(f"Total de faltantes: {len(faltantes)}\n")
        arq.write("=" * 80 + "\n\n")

        for (codigo_emp, nome_emp), funcs in sorted(por_empresa.items()):
            arq.write(f"EMPRESA: {nome_emp} (Código: {codigo_emp})\n")
            arq.write(f"Faltantes: {len(funcs)}\n")
            arq.write("-" * 60 + "\n")
            for f in funcs:
                arq.write(f"  - {f['funcionario_nome']}  [{f['motivo']}]\n")
            arq.write("\n")

    logger.info(f"Relatório salvo em: {caminho_saida}")


# ============================================================
# EXECUÇÃO PRINCIPAL COM INTERFACE
# ============================================================

def executar_analise_completa(pdf_geral_path, pasta_individual_path, callback_progresso=None, competencia=None):
    """
    Executa a análise completa usando os caminhos fornecidos pela interface.
    competencia: string 'MM/AAAA' — funcionários admitidos após essa competência são ignorados.
    Retorna um resumo do resultado.
    """
    if callback_progresso:
        callback_progresso(0, "Iniciando análise...")

    logger.info("Iniciando análise GMS... Competência: %s", competencia or "não informada")

    # 1. Mapear arquivos individuais de GMS
    if callback_progresso:
        callback_progresso(10, "Mapeando arquivos individuais...")
    mapa_arquivos = mapear_arquivos_individuais(pasta_individual_path)

    # 2. Extrair funcionários do relatório geral de empregados
    if callback_progresso:
        comp_info = f" (competência {competencia})" if competencia else ""
        callback_progresso(30, f"Lendo empregados: {os.path.basename(pdf_geral_path)}{comp_info}")
    logger.info(f"Lendo empregados: {pdf_geral_path}")
    funcionarios_geral = extrair_funcionarios_empregados(pdf_geral_path, competencia=competencia)

    if not funcionarios_geral:
        raise Exception("Não foi possível extrair funcionários do relatório de empregados.")

    total_empresas = len(funcionarios_geral)
    total_funcionarios = sum(len(v) for v in funcionarios_geral.values())
    logger.info(f"Empregados: {total_empresas} empresas, {total_funcionarios} funcionários")

    if callback_progresso:
        callback_progresso(60, f"Encontrados {total_empresas} empresas e {total_funcionarios} funcionários")

    # 3. Comparar
    if callback_progresso:
        callback_progresso(80, "Comparando funcionários...")
    faltantes = comparar_funcionarios(funcionarios_geral, mapa_arquivos)

    # 4. Gerar relatório
    if callback_progresso:
        callback_progresso(90, "Gerando relatório...")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    pasta_faltantes = os.path.join(script_dir, "faltantes")
    os.makedirs(pasta_faltantes, exist_ok=True)
    output_path = os.path.join(pasta_faltantes, f"faltantes_{timestamp}.txt")

    if faltantes:
        gerar_relatorio_faltantes(faltantes, output_path)
        logger.info(f"{len(faltantes)} funcionários faltantes -> {output_path}")

        resumo = f"Encontrados {len(faltantes)} funcionários faltantes em {total_empresas} empresas.\nRelatório salvo em: {os.path.basename(output_path)}"
        mensagem_notificacao = (
            f"⚠️ <@1285316758009544844> <@&1299044385899548752> Análise GMS concluída: {len(faltantes)} funcionário(s) faltante(s) em {total_empresas} empresa(s)."
            f"\nRelatório: {os.path.basename(output_path)}"
        )
    else:
        # Criar arquivo vazio indicando que não há faltantes
        with open(output_path, "w", encoding="utf-8") as arq:
            arq.write("=" * 80 + "\n")
            arq.write("RELATÓRIO DE FUNCIONÁRIOS FALTANTES NA GMS\n")
            arq.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            arq.write("Nenhum funcionário faltante encontrado!\n")
            arq.write("=" * 80 + "\n")

        logger.info("Nenhum funcionário faltante. Todos presentes nos GMS!")
        resumo = f"Nenhum funcionário faltante encontrado!\nTodas as {total_funcionarios} funcionários estão presentes nos relatórios individuais."
        mensagem_notificacao = (
            f"✅ <@1285316758009544844> Análise GMS concluída: nenhum funcionário faltante encontrado."
            f"\nTotal de funcionários verificados: {total_funcionarios}."
        )

    enviar_webhook_discord(mensagem_notificacao)

    if callback_progresso:
        callback_progresso(100, "Análise concluída!")

    return resumo


# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================

def main():
    # Verificar se foi chamado com argumentos de linha de comando
    if len(os.sys.argv) > 1:
        # Modo linha de comando
        if len(os.sys.argv) != 3:
            print("Uso: python Analytics_GMS.py <pdf_geral> <pasta_individual>")
            return

        pdf_geral = os.sys.argv[1]
        pasta_individual = os.sys.argv[2]

        try:
            resultado = executar_analise_completa(pdf_geral, pasta_individual)
            print("\n" + "="*50)
            print("RESULTADO DA ANÁLISE:")
            print("="*50)
            print(resultado)
        except Exception as e:
            print(f"Erro: {e}")
            return
    else:
        # Modo interface gráfica
        app = AnalyticsGUI()
        app.run()


if __name__ == "__main__":
    main()
