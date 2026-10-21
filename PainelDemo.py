import re
import sys

CRITERIOS = {
    "tipo_contratacao": {
        "Estágio": [r"\bestágio\b", r"\bestagio\b", r"\bintern\b", r"\binternship\b"],
        "Trainee": [r"\btrainee\b"],
        "Emprego (CLT)": [r"\bclt\b", r"\befetivo\b"],
        "Pessoa Jurídica (PJ)": [r"\bpj\b", r"\bpessoa jurídica\b", r"\bpessoa juridica\b", r"\bprestador de serviço\b"],
        "Temporário": [r"\btemporário\b", r"\btemporario\b", r"\btemporary\b"],
        "Meio período": [r"\bmeio período\b", r"\bmeio periodo\b", r"\bpart[- ]time\b"],
        "Voluntariado": [r"\bvolunt[aá]ri[ao]\b", r"\bvolunteer\b"]
    },
    "modelo_trabalho": {
        "Remoto": [r"\bremoto\b", r"\bremote\b", r"\bhome office\b", r"\b100% remoto\b"],
        "Híbrido": [r"\bhíbrido\b", r"\bhibrido\b", r"\bhybrid\b"],
        "No local": [r"\bno local\b", r"\bpresencial\b", r"\bon-site\b", r"\bonsite\b"]
    },
    "nivel_vaga": {
        "Early Career": [r"\bestágio\b", r"\bestagio\b", r"\btrainee\b", r"\bjúnior\b", r"\bjunior\b", r"\bjr\b", r"\bentry[- ]level\b"],
        "Professional/Experienced": [r"\bpleno\b", r"\bpl\b", r"\bsênior\b", r"\bsenior\b", r"\bsr\b", r"\bmid[- ]level\b"],
        "Advanced/Expert": [r"\bstaff\b", r"\bprincipal\b", r"\bespecialista\b", r"\bspecialist\b", r"\bexpert\b"],
        "Leadership": [r"\bmanager\b", r"\bgerente\b", r"\bgestor\b", r"\bhead\b", r"\bdiretor\b", r"\bdirector\b", r"\bcoordenador\b", r"\blead\b"]
    },
    "clusters": {
        "Cyber": [r"seguran[çc]a cibern[eé]tica", r"cyber\s*security", r"defensive security", r"offensive security", r"appsec", r"infosec"],
        "IA": [r"intelig[eê]ncia artificial", r"\bia\b", r"\bai\b", r"machine learning", r"agentes de ia", r"ai scientist", r"multi[- ]agents?"],
        "Dados": [r"data science", r"ci[eê]ncia de dados", r"data analytics", r"engenharia de dados", r"data engineering", r"business intelligence", r"\bbi\b"],
        "Dev": [r"an[aá]lise e desenvolvimento de sistemas", r"engenharia de software", r"software architecture", r"java", r"\.net", r"full\s*stack", r"backend", r"desenvolvedor", r"developer"],
        "Cloud & Infraestrutura": [r"computa[çc][aã]o em nuvem", r"cloud", r"devops", r"aws", r"azure", r"gcp", r"infraestrutura"],
        "Creative Tech": [r"design digital", r"web design", r"game development", r"front\s*end", r"ui/ux", r"ux design"],
        "Robotics & Systems": [r"mecatr[oô]nica", r"engenharia de computa[çc][aã]o", r"sistemas embarcados", r"rob[oó]tica"],
        "Tech Business & Liderança": [r"gest[aã]o de ti", r"tech management", r"product manager", r"gest[aã]o estrat[eé]gica", r"project management", r"cto", r"cio", r"scrum master"]
    }
}

PADROES_AFIRMATIVA = [
    r"\bpcd\b",
    r"pessoa[s]? com defici[eê]ncia",
    r"\bmulher(es)?\b",
    r"preto[s]?",
    r"preta[s]?",
    r"pardo[s]?",
    r"parda[s]?",
    r"negro[s]?",
    r"negra[s]?",
    r"afirmativ[ao]",
    r"diversidade",
    r"inclus[aã]o"
]

def verificar_afirmativa(texto: str) -> bool:
    texto_lower = texto.lower()
    return any(re.search(padrao, texto_lower) for padrao in PADROES_AFIRMATIVA)

def extrair_titulo(texto: str, eh_afirmativa: bool) -> str:
    linhas = [linha.strip() for linha in texto.strip().splitlines() if linha.strip()]
    titulo_base = linhas[0] if linhas else "Vaga sem título informado"
    
    # Remove prefixos comuns como "Título:", "Vaga:", etc.
    titulo_base = re.sub(r"^(t[ií]tulo da vaga|t[ií]tulo|vaga|cargo):\s*", "", titulo_base, flags=re.IGNORECASE)
    
    if eh_afirmativa:
        if not re.search(r"afirmativ[ao]", titulo_base, re.IGNORECASE):
            return f"{titulo_base} - Vaga Afirmativa"
    return titulo_base

def gerar_resumo(texto: str, limite: int = 2000) -> str:
    texto_limpo = re.sub(r"\s+", " ", texto).strip()
    if len(texto_limpo) <= limite:
        return texto_limpo
    return texto_limpo[:limite - 3] + "..."

def classificar_vaga(texto: str):
    texto_lower = texto.lower()
    eh_afirmativa = verificar_afirmativa(texto)
    
    resultado = {
        "Título": extrair_titulo(texto, eh_afirmativa),
        "Tipo de contratação": set(),
        "Modelo de trabalho": set(),
        "Nível da vaga": set(),
        "Clusters identificados": set(),
        "Resumo": gerar_resumo(texto, limite=2000)
    }

    for categoria, padroes in CRITERIOS["tipo_contratacao"].items():
        if any(re.search(p, texto_lower) for p in padroes):
            resultado["Tipo de contratação"].add(categoria)

    for categoria, padroes in CRITERIOS["modelo_trabalho"].items():
        if any(re.search(p, texto_lower) for p in padroes):
            resultado["Modelo de trabalho"].add(categoria)

    for nivel, padroes in CRITERIOS["nivel_vaga"].items():
        if any(re.search(p, texto_lower) for p in padroes):
            resultado["Nível da vaga"].add(nivel)

    for cluster, padroes in CRITERIOS["clusters"].items():
        if any(re.search(p, texto_lower) for p in padroes):
            resultado["Clusters identificados"].add(cluster)

    return resultado

def exibir_relatorio(resultado: dict):
    print("\n" + "=" * 60)
    print("                RELATÓRIO DA VAGA                ")
    print("=" * 60)
    print(f"Título                  : {resultado['Título']}")
    print(f"Tipo de contratação     : {', '.join(resultado['Tipo de contratação']) or 'Não identificado'}")
    print(f"Modelo de trabalho      : {', '.join(resultado['Modelo de trabalho']) or 'Não identificado'}")
    print(f"Nível da vaga           : {', '.join(resultado['Nível da vaga']) or 'Não identificado'}")
    print(f"Clusters identificados  : {', '.join(resultado['Clusters identificados']) or 'Não identificado'}")
    print("-" * 60)
    print("Resumo da vaga (até 2000 caracteres):")
    print(resultado["Resumo"])
    print("=" * 60 + "\n")

if __name__ == "__main__":
    print("Cole a descrição da vaga abaixo (Pressione Ctrl+D no Linux/Mac ou Ctrl+Z + Enter no Windows para finalizar):")
    texto_vaga = sys.stdin.read()
    
    if texto_vaga.strip():
        analise = classificar_vaga(texto_vaga)
        exibir_relatorio(analise)
    else:
        print("Nenhum texto informado.")
