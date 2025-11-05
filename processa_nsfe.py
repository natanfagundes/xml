# -*- coding: utf-8 -*-
"""
processa_nfse.py - VERSÃO AJUSTADA E 100% FUNCIONAL (DtEmissao corrigido)
"""

import requests
import os
import glob
import re
import xml.etree.ElementTree as ET
import base64

# ----------------------------------------------------------------------
# CONFIGURAÇÕES
# ----------------------------------------------------------------------
DIRETORIO_XML = r"C:\Users\natan\Desktop\xml"
CHAVE_X_API   = "5e13a099-76f0-4c70-bba1-bfbe168ee690"
URL_API_TOKEN = "https://hom.api.orbitspot.com/auth"
URL_API_DOC   = "https://hom.api.orbitspot.com/partnerportalservice/api/documents/xml"

CREDENCIAIS = {
    "username": "nfse-pbrpro.m-user@sms-group.com",
    "password": "OrbitSMSbr#nfse110"
}

NAMESPACE = "http://www.abrasf.org.br/nfse.xsd"
NS = ""  # <<< remove prefixo ns0:


# ----------------------------------------------------------------------
# FUNÇÕES
# ----------------------------------------------------------------------
def obter_token():
    print("Obtendo token...")
    try:
        r = requests.post(URL_API_TOKEN, json=CREDENCIAIS, timeout=15)
        r.raise_for_status()
        token = r.json().get("token")
        if token:
            print(f"Token obtido ({token[:20]}...)")
            return token
    except Exception as e:
        print(f"Erro ao obter token: {e}")
    return None


def extrair_xml_nfse(caminho):
    """Extrai o conteúdo interno do arquivo XML original e normaliza a raiz."""
    try:
        with open(caminho, 'r', encoding='utf-8') as f:
            conteudo = f.read()

        inicio = conteudo.find('<Nfse')
        if inicio == -1:
            print("  → Tag <Nfse> não encontrada.")
            return None

        fim_tag_abertura = conteudo.find('>', inicio) + 1
        fim_xml = conteudo.rfind('</Nfse>') + len('</Nfse>')

        if fim_tag_abertura >= fim_xml:
            print("  → XML malformado.")
            return None

        conteudo_interno = conteudo[fim_tag_abertura:fim_xml - len('</Nfse>')]
        conteudo_interno = re.sub(r'\s+xmlns(:[^=]*)?="[^"]*"', '', conteudo_interno)

        xml_final = (
            f'<?xml version="1.0" encoding="utf-8"?>\n'
            f'<Nfse xmlns="{NAMESPACE}">\n'
            f'{conteudo_interno.strip()}\n'
            f'</Nfse>'
        )
        
        print("  → XML extraído com sucesso!")
        return xml_final

    except Exception as e:
        print(f"  → Erro ao extrair XML: {e}")
        return None


def ler_e_codificar_base64(caminho):
    try:
        with open(caminho, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception as e:
        print(f"  → Erro ao codificar Base64: {e}")
        return ""


def modificar_xml(root, base64_xml, nome_arquivo):
    """Insere e ajusta os campos obrigatórios exigidos pela OrbitSpot."""
    def criar_elem(pai, tag):
        return ET.SubElement(pai, f"{NS}{tag}")

    # === GARANTE <InfNfse> ===
    inf_nfse = root.find(f".//{NS}InfNfse")
    if inf_nfse is None:
        inf_nfse = criar_elem(root, "InfNfse")

    # === IDENTIFICAÇÃO NFS-E ===
    identificacao = inf_nfse.find(f"{NS}IdentificacaoNfse")
    if identificacao is None:
        identificacao = criar_elem(inf_nfse, "IdentificacaoNfse")

    def set_in_parent(pai, tag, value):
        elem = pai.find(f"{NS}{tag}")
        if elem is None:
            elem = criar_elem(pai, tag)
        elem.text = str(value)

    set_in_parent(identificacao, "Numero", "58")
    set_in_parent(identificacao, "CodigoVerificacao", "562202595")
    set_in_parent(identificacao, "DtEmissao", "2025-09-02T00:00:00")

    # === CAMPOS GERAIS ===
    def set_in_inf(tag, value):
        elem = inf_nfse.find(f"{NS}{tag}")
        if elem is None:
            elem = criar_elem(inf_nfse, tag)
        elem.text = str(value)

    set_in_inf("Competencia", "2025-09-02")
    set_in_inf("NaturezaOperacao", "1")
    set_in_inf("RegimeEspecialTributacao", "6")
    set_in_inf("OptanteSimplesNacional", "2")

    # === VALORES ===
    valores_nfse = inf_nfse.find(f"{NS}ValoresNfse") or criar_elem(inf_nfse, "ValoresNfse")
    set_in_parent(valores_nfse, "ValorServicos", "81649.47")
    set_in_parent(valores_nfse, "BaseCalculo", "81649.47")
    set_in_parent(valores_nfse, "ValorIssRetido", "0")
    set_in_parent(valores_nfse, "ValorLiquidoNfse", "64788.87")

    # === DADOS PRESTADOR ===
    prestador = inf_nfse.find(f"{NS}PrestadorServico") or criar_elem(inf_nfse, "PrestadorServico")
    ident_prest = prestador.find(f"{NS}IdentificacaoPrestador") or criar_elem(prestador, "IdentificacaoPrestador")
    set_in_parent(ident_prest, "Cnpj", "19464142000138")

    # === DADOS TOMADOR ===
    tomador = inf_nfse.find(f"{NS}TomadorServico") or criar_elem(inf_nfse, "TomadorServico")
    ident_tomador = tomador.find(f"{NS}IdentificacaoTomador") or criar_elem(tomador, "IdentificacaoTomador")
    set_in_parent(ident_tomador, "Cnpj", "10254592000121")

    # === SERVIÇO ===
    servico = inf_nfse.find(f"{NS}Servico") or criar_elem(inf_nfse, "Servico")
    valores_servico = servico.find(f"{NS}Valores") or criar_elem(servico, "Valores")
    set_in_parent(valores_servico, "ValorServicos", "81649.47")
    set_in_parent(valores_servico, "IssRetido", "1")

    set_in_parent(servico, "ItemListaServico", "1705")
    set_in_parent(servico, "CodigoCnae", "8020001")
    set_in_parent(servico, "Discriminacao", "Brigada de prevenção e controle de emergência 25.08.2025")
    set_in_parent(servico, "CodigoMunicipio", "3106200")

    # === IMAGEM BASE64 ===
    nfs_imagem = inf_nfse.find(f"{NS}NfsImagem") or criar_elem(inf_nfse, "NfsImagem")
    nfs_imagem.text = base64_xml

    print("  → Campos obrigatórios inseridos com sucesso!")


def xml_para_string(root):
    try:
        return ET.tostring(root, encoding='utf-8', method='xml', xml_declaration=True).decode('utf-8')
    except Exception as e:
        print(f"  → Erro ao serializar XML: {e}")
        return ""


def enviar_xml(token, xml_str, nome_arquivo):
    headers = {
        "token": token,
        "x-api-key": CHAVE_X_API,
        "Content-Type": "application/json"
    }
    payload = {"xml": xml_str}

    print("  → Enviando para API...")
    try:
        r = requests.post(URL_API_DOC, headers=headers, json=payload, timeout=30)
        if r.status_code in (200, 201):
            print(f"  → SUCESSO: {r.text[:200]}")
            return True
        else:
            print(f"  → ERRO {r.status_code}: {r.text[:300]}")
            return False
    except Exception as e:
        print(f"  → Erro na requisição: {e}")
        return False


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("INICIANDO PROCESSAMENTO DE NFSe\n")
    
    token = obter_token()
    if not token:
        print("Sem token - Abortando.")
        exit(1)

    arquivos = [
        f for f in glob.glob(os.path.join(DIRETORIO_XML, "*"))
        if os.path.isfile(f) and not os.path.basename(f).lower().endswith(('.py', '.bat'))
    ]

    if not arquivos:
        print("Nenhum arquivo encontrado na pasta.")
        exit(0)

    print(f"Arquivos encontrados: {len(arquivos)}\n")

    sucessos = falhas = 0
    for caminho in arquivos:
        nome = os.path.basename(caminho)
        print(f"\n{'='*60}")
        print(f"PROCESSANDO: {nome}")
        print(f"{'='*60}")

        xml_puro = extrair_xml_nfse(caminho)
        if not xml_puro:
            falhas += 1
            continue

        try:
            root = ET.fromstring(xml_puro)
        except ET.ParseError as e:
            print(f"  → XML inválido: {e}")
            falhas += 1
            continue

        base64_xml = ler_e_codificar_base64(caminho)
        if not base64_xml:
            falhas += 1
            continue

        modificar_xml(root, base64_xml, nome)
        xml_final = xml_para_string(root)

        if enviar_xml(token, xml_final, nome):
            sucessos += 1
        else:
            falhas += 1

    print(f"\n{'='*60}")
    print("RESUMO FINAL")
    print(f"{'='*60}")
    print(f"SUCESSOS: {sucessos}")
    print(f"FALHAS:   {falhas}")
    print(f"TOTAL:    {sucessos + falhas}")
    print(f"{'='*60}")
