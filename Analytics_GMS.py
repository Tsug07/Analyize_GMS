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
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("analytics.log", encoding="utf-8"),
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
        pasta_frame.pack(fill="x", padx=20, pady=(5, 20))

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
                      command=self.testar_webhook).pack(side="left")

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

    def executar_analise(self):
        if not os.path.exists(self.pdf_geral_path):
            messagebox.showerror("Erro", "O arquivo PDF geral selecionado não existe!")
            return

        if not os.path.exists(self.pasta_individual_path):
            messagebox.showerror("Erro", "A pasta selecionada não existe!")
            return

        try:
            self.status_label.configure(text="Iniciando análise...")
            self.progress.set(0)
            self.root.update()

            resultado = executar_analise_completa(
                self.pdf_geral_path,
                self.pasta_individual_path,
                self.atualizar_progresso
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
        'JUNIOR', 'LICITAÇÃO', 'LIMPEZA', 'LINHA', 'LOGISTICA', 'LOJA',
        'MAQUINAS', 'MECANICO', 'MARKETING', 'MONTADOR', 'MONITORAMENTO', 'MOTORISTA',
        'OPERACIONAL', 'OPERADOR', 'OPERADORA', 'PESSOAL', 'PLANEJAMENTO',
        'PLENO', 'PRODUCAO', 'PRODUÇÃO', 'PROGRAMADOR', 'PROMOTOR',
        'RECEPCIONISTA', 'REFRIGERACAO', 'REFRIGERAÇÃO', 'REPOSITOR', 'RESPONSAVEL',
        'SECRETARIA', 'SECCIONADORA', 'SENIOR', 'SERVIÇOS', 'SERVICOS',
        'SUPERVISOR', 'SUPERVISORA', 'SUPORTE', 'SUPRIMENTOS',
        'TECNICO', 'TÉCNICO', 'VENDAS', 'VENDEDOR', 'VENDEDORA',
        'ACOUGUEIRO', 'BICICLET', 'CFTV',
        'OP', 'DE', 'DO', 'DA', 'EM', 'E', 'I', 'II', 'III', 'IV', 'IX',
        'JR', 'PL', 'SR', 'ADM', 'COM', 'PÓS', 'TI', 'RH',
        '(A)', '(FARMACIA)',
    }

    palavras = trecho.split()
    if not palavras:
        return ""

    # Percorre de trás para frente enquanto a palavra parecer cargo
    idx_corte = len(palavras)
    for i in range(len(palavras) - 1, -1, -1):
        palavra = palavras[i].upper().strip('.,()')
        if palavra in palavras_cargo or palavra.startswith('AUX'):
            idx_corte = i
        else:
            break

    # O nome são as palavras antes do cargo
    nome = ' '.join(palavras[:idx_corte])
    return nome.strip()


def extrair_funcionarios_empregados(caminho_pdf: str) -> dict:
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
            # Detectar início da lista (linha do cabeçalho da tabela)
            if "CódigoNome" in linha.replace(" ", "") or ("Código" in linha and "Nome" in linha and "Cargo" in linha):
                em_lista = True
                continue

            # Detectar fim da lista
            if "Total de empregados" in linha:
                em_lista = False
                continue

            if not em_lista:
                continue

            # Linhas de funcionário começam com código numérico colado ao nome
            # Formato: "5037ELENI TEREZA RODRIGUES OPERADORA DE CAIXA Mensalista 180,00 1101/02/2019 NS"
            # Estratégia: capturar tudo entre o código e "Mensalista|Horista|..." ,
            # depois extrair o nome usando o padrão numérico "DDD,DD" (horas) como âncora reversa.
            match_categoria = re.match(
                r'^(\d+)(.+?)\s+(Mensalista|Horista|Diarista|Tarefeiro)\s+(\d+[,.]\d{2})',
                linha, re.IGNORECASE
            )
            if match_categoria:
                # O trecho entre código e categoria contém "NOME CARGO"
                # O cargo vem logo antes da categoria. Usamos o valor de horas (ex: 180,00)
                # para separar: tudo antes do cargo é o nome.
                # Como não sabemos onde o nome termina e o cargo começa,
                # usamos heurística: o nome é a sequência mais longa de palavras
                # totalmente em maiúsculas desde o início, até encontrar uma palavra
                # que faça parte de um cargo comum.
                trecho = match_categoria.group(2).strip()
                nome_func = _extrair_nome_do_trecho(trecho)
                if nome_func:
                    funcionarios.append(normalizar_nome(nome_func))

        if funcionarios:
            if nome_empresa in empresas:
                # Empresa com múltiplas páginas
                empresas[nome_empresa].extend(funcionarios)
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
        if not arquivo.startswith("Relatorio_") or not arquivo.endswith(".pdf"):
            continue

        # Extrair código e nome: Relatorio_[ID]_[NOME].pdf
        partes = arquivo.replace(".pdf", "").split("_", 2)
        if len(partes) < 3:
            logger.warning(f"Arquivo com formato inesperado: {arquivo}")
            continue

        codigo = partes[1]
        nome_empresa = partes[2].replace("_", " ")

        mapa[nome_empresa] = {
            "arquivo": arquivo,
            "caminho": os.path.join(diretorio, arquivo),
            "codigo": codigo,
            "nome_empresa": nome_empresa
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

def executar_analise_completa(pdf_geral_path, pasta_individual_path, callback_progresso=None):
    """
    Executa a análise completa usando os caminhos fornecidos pela interface.
    Retorna um resumo do resultado.
    """
    if callback_progresso:
        callback_progresso(0, "Iniciando análise...")

    logger.info("Iniciando análise GMS...")

    # 1. Mapear arquivos individuais de GMS
    if callback_progresso:
        callback_progresso(10, "Mapeando arquivos individuais...")
    mapa_arquivos = mapear_arquivos_individuais(pasta_individual_path)

    # 2. Extrair funcionários do relatório geral de empregados
    if callback_progresso:
        callback_progresso(30, f"Lendo empregados: {os.path.basename(pdf_geral_path)}")
    logger.info(f"Lendo empregados: {pdf_geral_path}")
    funcionarios_geral = extrair_funcionarios_empregados(pdf_geral_path)

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

    # Criar nome do arquivo de saída na pasta GMS_Analitics (não misturar com os PDFs)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, f"faltantes_{timestamp}.txt")

    if faltantes:
        gerar_relatorio_faltantes(faltantes, output_path)
        logger.info(f"{len(faltantes)} funcionários faltantes -> {output_path}")

        resumo = f"Encontrados {len(faltantes)} funcionários faltantes em {total_empresas} empresas.\nRelatório salvo em: {os.path.basename(output_path)}"
        mensagem_notificacao = (
            f"⚠️ Análise GMS concluída: {len(faltantes)} funcionário(s) faltante(s) em {total_empresas} empresa(s)."
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
            f"✅ Análise GMS concluída: nenhum funcionário faltante encontrado."
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
