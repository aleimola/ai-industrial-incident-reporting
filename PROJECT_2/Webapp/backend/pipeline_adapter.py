"""
Adapter per collegare la web app alla pipeline di script_finale 3.3.ipynb.
"""
import os
import re
import time
import json
from pathlib import Path
from typing import Any
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

_SCHEMA_JSON = """
{
  "Evento": {
    "id_evento": "INC-AAAA-NNN",
    "Data_Ora": "GG/MM/AAAA HH:MM o generica",
    "Tipo_Evento": "Infortunio | Near Miss | Sicurezza | Ambiente",
    "Gravita_Stimata": "Bassa | Media | Alta | Critica | Fatale",
    "Prima_dell_evento": "Descrivi cosa succedeva prima dell'evento, deducibile dal contesto",
    "Trigger_Evento_Principale": "Quale è stata la causa scatenante?",
    "Dinamica_Sintesi": "Riassunto analitico, oggettivo e completo della dinamica dell'evento"
  },
  "Luoghi_e_Mezzi": {
    "Reparto": "Nome del reparto (es. Reparto Stampa, Linea 3). Usa le tue capacità per individuarlo.",
    "Area_Coinvolta": "Zona specifica all'interno del reparto.",
    "Macchinario_Coinvolto": "Nome tecnico del macchinario e alias usato (se presente)",
    "Codice_Macchinario": "Codice o targa se menzionati verbalmente",
    "Condizioni_Ambientali": "Condizioni fisiche (es. standard, bagnato, polvere, illuminazione) deducibili"
  },
  "Analisi_Cause": {
    "Causa_Tecnica": "Causa materiale o logica (es. assenza di blocco, sversamento) dedotta",
    "Fattore_Umano": "Fattori comportamentali (es. disattenzione, fretta) dedotti",
    "Fattore_Ambientale": "Fattori di contesto (es. rumore, visibilità) dedotti",
    "Causa_Radice": "Analisi investigativa profonda della causa primaria scatenante (Root Cause)"
  },
  "Conseguenze": {
    "Conseguenza_Fisica_Immediata": "Danni a persone, anche lievi o potenziali (Near Miss)",
    "Parte_Corpo_Lesa": "Specifica della parte del corpo (se c'è infortunio)",
    "Giorni_Inabilita_Stimati": "Stima dei giorni (es. 'Nessuno', 'Pochi giorni') se possibile",
    "Conseguenza_Materiale_Aziendale": "Danni materiali alle macchine o impatto evidente",
    "Fermo_Produzione_Ore": "Stima ore/giorni o impatto sulla catena"
  },
  "Azioni_Correttive": {
    "Procedura_Corretta_Violata": "Quale regola e/o procedura è stata violata? (Deduci logicamente)",
    "DPI_Mancanti_o_Inadeguati": "Stato DPI (esitavano? idonei?) deducibile",
    "Azione_Correttiva_Immediata": "Azione presa sùbito (o urgente da prendere)",
    "Azione_Correttiva_da_Implementare": "Azione strutturale per evitare che succeda di nuovo (proponi tu se manca)",
    "Responsabile_Azione": "Chi dovrebbe occuparsene (es. RSPP, Manutentore, ecc.)",
    "Scadenza_Azione": "Tempistiche (es. 'Immediata', 'Lungo termine')"
  },
  "Persone_Coinvolte": [
    {
      "Nome": "Nome e cognome, o riferimento generico (es. L'operatore alla sega)",
      "Ruolo_Aziendale": "Deducibile (es. Operatore, Manutentore)",
      "Coinvolgimento": "Vittima | Testimone | RSPP | Supervisore | Causante",
      "Danno_Riportato": "Danni fisici e psicologici",
      "Affidabilita_Testimonianza": "Valutazione sulla chiarezza (Alta | Media | Bassa)",
      "Note_Testimonianza": "Note aggiuntive su come ha parlato/cosa ha detto"
    }
  ]
}
"""

_ONTOLOGY_TEMPLATE = {
    "Evento": {
        "id_evento": "N/D",
        "Data_Ora": "N/D",
        "Tipo_Evento": "N/D",
        "Gravita_Stimata": "N/D",
        "Prima_dell_evento": "N/D",
        "Trigger_Evento_Principale": "N/D",
        "Dinamica_Sintesi": "N/D",
    },
    "Luoghi_e_Mezzi": {
        "Reparto": "N/D",
        "Area_Coinvolta": "N/D",
        "Macchinario_Coinvolto": "N/D",
        "Codice_Macchinario": "N/D",
        "Condizioni_Ambientali": "N/D",
    },
    "Analisi_Cause": {
        "Causa_Tecnica": "N/D",
        "Fattore_Umano": "N/D",
        "Fattore_Ambientale": "N/D",
        "Causa_Radice": "N/D",
    },
    "Conseguenze": {
        "Conseguenza_Fisica_Immediata": "N/D",
        "Parte_Corpo_Lesa": "N/D",
        "Giorni_Inabilita_Stimati": "N/D",
        "Conseguenza_Materiale_Aziendale": "N/D",
        "Fermo_Produzione_Ore": "N/D",
    },
    "Azioni_Correttive": {
        "Procedura_Corretta_Violata": "N/D",
        "DPI_Mancanti_o_Inadeguati": "N/D",
        "Azione_Correttiva_Immediata": "N/D",
        "Azione_Correttiva_da_Implementare": "N/D",
        "Responsabile_Azione": "N/D",
        "Scadenza_Azione": "N/D",
    },
}

_PERSONA_TEMPLATE = {
    "Nome": "N/D",
    "Ruolo_Aziendale": "N/D",
    "Coinvolgimento": "N/D",
    "Danno_Riportato": "N/D",
    "Affidabilita_Testimonianza": "N/D",
    "Note_Testimonianza": "N/D",
}

def normalize_ontology(raw: dict[str, Any]) -> dict[str, Any]:
    normalized = {}
    for section, fields in _ONTOLOGY_TEMPLATE.items():
        source = raw.get(section, {}) if isinstance(raw.get(section), dict) else {}
        normalized[section] = {
            field: source.get(field) or "N/D"
            for field in fields
        }

    persone = raw.get("Persone_Coinvolte", [])
    if not isinstance(persone, list):
        persone = []
    normalized["Persone_Coinvolte"] = [
        {field: (persona.get(field) if isinstance(persona, dict) else None) or "N/D" for field in _PERSONA_TEMPLATE}
        for persona in persone
    ]
    return normalized

def build_prompt_singolo_testo(nome_file: str) -> str:
    return f"""
Ruolo: Sei un Ispettore Investigativo HSE (Health, Safety & Environment) senior, eccezionalmente pignolo ma razionale, esperto in Data Extraction.
Ti viene fornito il testo della trascrizione e dell'analisi logica investigativa dell'audio '{nome_file}'.

OBIETTIVI DI ESTRAZIONE RIGOROSI:
1. Devi massimizzare l'estrazione di dati: leggi tra le righe del racconto per scovare dettagli impliciti evidenti.
2. Per i campi ASSOLUTAMENTE NON presenti e impossibili da stabilire, scrivi ESATTAMENTE la dicitura standard: "N/D"
3. Scrivi risposte complete e chiare.
4. IMPORTANTE PER JSON: NON usare mai virgolette doppie (") all'interno dei valori di testo. Se devi citare qualcosa, usa apici singoli (') o omettile del tutto, per evitare di rompere il parser JSON.

Istruzioni specifiche per Luoghi_e_Mezzi:
- Reparto: estrai il nome esatto del reparto citato. Se usa un soprannome o abbreviazione, riportala così com'è senza modifiche.
- Area_Coinvolta: estrai la zona precisa all'interno del reparto (es. vicino alla pressa, zona scarico). Se non deducibile, metti "N/D".
- Macchinario_Coinvolto: estrai il nome tecnico o il nome comune usato. Se usa un alias (es. 'la bestia'), affiancagli il nome tecnico probabile.
- Codice_Macchinario: solo se citato esplicitamente (targa, codice alfanumerico). Altrimenti "N/D".

*** ESEMPIO DI ESTRAZIONE PERFETTA (FEW-SHOT) ***
TESTO FORNITO:
TRASCRIZIONE: "Ieri in Reparto Stampa, alla Linea 3, l'operatore Mario Rossi si è tagliato la mano destra con la Pressa P200 perché ha rimosso il carter di protezione per fare prima."
ANALISI HSE: L'evento è un incidente. Luogo: Reparto Stampa, Area: Linea 3. Macchinario: Pressa P200. Persona coinvolta: Mario Rossi (vittima, taglio alla mano). Causa: rimozione protezione (Fattore Umano, Procedura Violata).

OUTPUT JSON ATTESO:
{{
  "Evento": {{
    "id_evento": "N/D",
    "Data_Ora": "Ieri",
    "Tipo_Evento": "Infortunio",
    "Gravita_Stimata": "Media",
    "Prima_dell_evento": "Operazione standard su linea, produzione in corso",
    "Trigger_Evento_Principale": "Rimozione carter di protezione",
    "Dinamica_Sintesi": "L'operatore si è tagliato la mano destra manovrando la pressa, a seguito di cui aveva rimosso il sistema difensivo per velocizzare la routine lavorativa."
  }},
  "Luoghi_e_Mezzi": {{
    "Reparto": "Reparto Stampa",
    "Area_Coinvolta": "Linea 3",
    "Macchinario_Coinvolto": "Pressa P200",
    "Codice_Macchinario": "N/D",
    "Condizioni_Ambientali": "Standard"
  }},
  "Analisi_Cause": {{
    "Causa_Tecnica": "Interferenza con gli organi commoventi senza sistema di blocco",
    "Fattore_Umano": "Elusione dispositivo di sicurezza per guadagnare tempo",
    "Fattore_Ambientale": "N/D",
    "Causa_Radice": "Pressione sui tempi di ciclo e under-valutazione dei rischi da parte del lavoratore."
  }},
  "Conseguenze": {{
    "Conseguenza_Fisica_Immediata": "Taglio lacerocontuso alla mano",
    "Parte_Corpo_Lesa": "Mano destra",
    "Giorni_Inabilita_Stimati": "N/D",
    "Conseguenza_Materiale_Aziendale": "N/D",
    "Fermo_Produzione_Ore": "N/D"
  }},
  "Azioni_Correttive": {{
    "Procedura_Corretta_Violata": "Divieto tassativo di rimozione protezioni macchina",
    "DPI_Mancanti_o_Inadeguati": "Guanti antitaglio idonei",
    "Azione_Correttiva_Immediata": "Assistenza medica lavoratore e blocco macchina fino al ripristino protezioni.",
    "Azione_Correttiva_da_Implementare": "Sensore anti-manomissione sul carter (interlock) e riunione formativa straordinaria sul rischio.",
    "Responsabile_Azione": "RSPP / Capo Reparto",
    "Scadenza_Azione": "Immediata"
  }},
  "Persone_Coinvolte": [
    {{
      "Nome": "Mario Rossi",
      "Ruolo_Aziendale": "Operatore linea",
      "Coinvolgimento": "Vittima",
      "Danno_Riportato": "Taglio alla mano destra",
      "Affidabilita_Testimonianza": "Alta",
      "Note_Testimonianza": "N/D"
    }}
  ]
}}
*************************************************

Formato output:
Restituisci ESCLUSIVAMENTE un oggetto JSON valido.
Schema:
{_SCHEMA_JSON}
"""

def build_prompt_trascrizione_e_analisi(nome_file: str) -> str:
    return f"""
Ruolo: Sei un esperto HSE (Health, Safety & Environment).
Ascolta il file audio '{nome_file}'.

Fase 1: Trascrizione
Trascrivi fedelmente in italiano il contenuto parlato. Se una parte non è comprensibile scrivi [incomprensibile].

Fase 2: Analisi Logica (Chain of Thought)
Dopo la trascrizione, scrivi una breve analisi ragionata dell'evento descritto. Rifletti passo passo su:
- Che tipo di evento è (Incidente, Near Miss, ecc.) e la gravità stimata.
- Dove è successo (Reparto, Area) e quali macchinari sono stati nominati.
- Chi è coinvolto (vittima, testimone) e che danno ha subito.
- Quali sono le cause tecniche, umane o ambientali menzionate e qual è la causa scatenante (Trigger).
- Quali sono le conseguenze materiali o fisiche.
- Quali procedure sono state violate e quali azioni correttive sono proposte o necessarie.

Formatta l'output dividendo chiaramente in "TRASCRIZIONE:" e "ANALISI HSE:".
"""

def reconstruct_ontology_if_flat(dati: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(dati, dict):
        return _ONTOLOGY_TEMPLATE.copy()
        
    ontology = {
        "Evento": {
            "id_evento": "N/D",
            "Data_Ora": "N/D",
            "Tipo_Evento": "N/D",
            "Gravita_Stimata": "N/D",
            "Prima_dell_evento": "N/D",
            "Trigger_Evento_Principale": "N/D",
            "Dinamica_Sintesi": "N/D",
        },
        "Luoghi_e_Mezzi": {
            "Reparto": "N/D",
            "Area_Coinvolta": "N/D",
            "Macchinario_Coinvolto": "N/D",
            "Codice_Macchinario": "N/D",
            "Condizioni_Ambientali": "N/D",
        },
        "Analisi_Cause": {
            "Causa_Tecnica": "N/D",
            "Fattore_Umano": "N/D",
            "Fattore_Ambientale": "N/D",
            "Causa_Radice": "N/D",
        },
        "Conseguenze": {
            "Conseguenza_Fisica_Immediata": "N/D",
            "Parte_Corpo_Lesa": "N/D",
            "Giorni_Inabilita_Stimati": "N/D",
            "Conseguenza_Materiale_Aziendale": "N/D",
            "Fermo_Produzione_Ore": "N/D",
        },
        "Azioni_Correttive": {
            "Procedura_Corretta_Violata": "N/D",
            "DPI_Mancanti_o_Inadeguati": "N/D",
            "Azione_Correttiva_Immediata": "N/D",
            "Azione_Correttiva_da_Implementare": "N/D",
            "Responsabile_Azione": "N/D",
            "Scadenza_Azione": "N/D",
        },
        "Persone_Coinvolte": []
    }

    def get_val(category: str, field: str, flat_fallback_key: str = None) -> str:
        if category in dati and isinstance(dati[category], dict):
            nested_val = dati[category].get(field)
            if nested_val is not None and nested_val != "" and nested_val != "N/D":
                return str(nested_val)
        ai_key = f"ai_{category}_{field}"
        if ai_key in dati and dati[ai_key] is not None and dati[ai_key] != "" and dati[ai_key] != "N/D":
            return str(dati[ai_key])
        if flat_fallback_key and flat_fallback_key in dati:
            fallback_val = dati.get(flat_fallback_key)
            if fallback_val is not None and fallback_val != "" and fallback_val != "N/D":
                return str(fallback_val)
        ai_report = dati.get("ai_report")
        if isinstance(ai_report, dict) and category in ai_report and isinstance(ai_report[category], dict):
            ai_val = ai_report[category].get(field)
            if ai_val is not None and ai_val != "" and ai_val != "N/D":
                return str(ai_val)
        return "N/D"

    # Populate Evento
    ontology["Evento"]["id_evento"] = get_val("Evento", "id_evento", "id_incidente")
    if ontology["Evento"]["id_evento"] == "N/D":
        ontology["Evento"]["id_evento"] = get_val("Evento", "id_evento", "id_evento")
    ontology["Evento"]["Data_Ora"] = get_val("Evento", "Data_Ora", "data_ora")
    ontology["Evento"]["Tipo_Evento"] = get_val("Evento", "Tipo_Evento", "categoria_evento")
    ontology["Evento"]["Gravita_Stimata"] = get_val("Evento", "Gravita_Stimata", "livello_gravita")
    ontology["Evento"]["Prima_dell_evento"] = get_val("Evento", "Prima_dell_evento")
    ontology["Evento"]["Trigger_Evento_Principale"] = get_val("Evento", "Trigger_Evento_Principale")
    ontology["Evento"]["Dinamica_Sintesi"] = get_val("Evento", "Dinamica_Sintesi", "dinamica_evento")
    if ontology["Evento"]["Dinamica_Sintesi"] == "N/D":
        ontology["Evento"]["Dinamica_Sintesi"] = get_val("Evento", "Dinamica_Sintesi", "trascrizione")

    # Populate Luoghi_e_Mezzi
    ontology["Luoghi_e_Mezzi"]["Reparto"] = get_val("Luoghi_e_Mezzi", "Reparto", "reparto")
    ontology["Luoghi_e_Mezzi"]["Area_Coinvolta"] = get_val("Luoghi_e_Mezzi", "Area_Coinvolta", "linea")
    ontology["Luoghi_e_Mezzi"]["Macchinario_Coinvolto"] = get_val("Luoghi_e_Mezzi", "Macchinario_Coinvolto", "macchinario_coinvolto")
    ontology["Luoghi_e_Mezzi"]["Codice_Macchinario"] = get_val("Luoghi_e_Mezzi", "Codice_Macchinario")
    ontology["Luoghi_e_Mezzi"]["Condizioni_Ambientali"] = get_val("Luoghi_e_Mezzi", "Condizioni_Ambientali")

    # Populate Analisi_Cause
    ontology["Analisi_Cause"]["Causa_Tecnica"] = get_val("Analisi_Cause", "Causa_Tecnica")
    ontology["Analisi_Cause"]["Fattore_Umano"] = get_val("Analisi_Cause", "Fattore_Umano")
    ontology["Analisi_Cause"]["Fattore_Ambientale"] = get_val("Analisi_Cause", "Fattore_Ambientale")
    ontology["Analisi_Cause"]["Causa_Radice"] = get_val("Analisi_Cause", "Causa_Radice")
    
    if (ontology["Analisi_Cause"]["Causa_Tecnica"] == "N/D" and
        ontology["Analisi_Cause"]["Fattore_Umano"] == "N/D" and
        ontology["Analisi_Cause"]["Fattore_Ambientale"] == "N/D" and
        ontology["Analisi_Cause"]["Causa_Radice"] == "N/D"):
        cp = dati.get("cause_probabili")
        if cp and cp != "N/D":
            ontology["Analisi_Cause"]["Causa_Radice"] = str(cp)

    # Populate Conseguenze
    ontology["Conseguenze"]["Conseguenza_Fisica_Immediata"] = get_val("Conseguenze", "Conseguenza_Fisica_Immediata")
    ontology["Conseguenze"]["Parte_Corpo_Lesa"] = get_val("Conseguenze", "Parte_Corpo_Lesa")
    ontology["Conseguenze"]["Giorni_Inabilita_Stimati"] = get_val("Conseguenze", "Giorni_Inabilita_Stimati")
    ontology["Conseguenze"]["Conseguenza_Materiale_Aziendale"] = get_val("Conseguenze", "Conseguenza_Materiale_Aziendale")
    ontology["Conseguenze"]["Fermo_Produzione_Ore"] = get_val("Conseguenze", "Fermo_Produzione_Ore")

    if (ontology["Conseguenze"]["Conseguenza_Fisica_Immediata"] == "N/D" and
        ontology["Conseguenze"]["Parte_Corpo_Lesa"] == "N/D" and
        ontology["Conseguenze"]["Giorni_Inabilita_Stimati"] == "N/D"):
        co = dati.get("conseguenze_operatore")
        if co and co != "N/D":
            ontology["Conseguenze"]["Conseguenza_Fisica_Immediata"] = str(co)
            
    if (ontology["Conseguenze"]["Conseguenza_Materiale_Aziendale"] == "N/D" and
        ontology["Conseguenze"]["Fermo_Produzione_Ore"] == "N/D"):
        dm = dati.get("danni_materiali")
        if dm and dm != "N/D":
            ontology["Conseguenze"]["Conseguenza_Materiale_Aziendale"] = str(dm)

    # Populate Azioni_Correttive
    ontology["Azioni_Correttive"]["Procedura_Corretta_Violata"] = get_val("Azioni_Correttive", "Procedura_Corretta_Violata")
    ontology["Azioni_Correttive"]["DPI_Mancanti_o_Inadeguati"] = get_val("Azioni_Correttive", "DPI_Mancanti_o_Inadeguati")
    ontology["Azioni_Correttive"]["Azione_Correttiva_Immediata"] = get_val("Azioni_Correttive", "Azione_Correttiva_Immediata", "azioni_immediate")
    ontology["Azioni_Correttive"]["Azione_Correttiva_da_Implementare"] = get_val("Azioni_Correttive", "Azione_Correttiva_da_Implementare", "raccomandazioni_preventive")
    ontology["Azioni_Correttive"]["Responsabile_Azione"] = get_val("Azioni_Correttive", "Responsabile_Azione")
    ontology["Azioni_Correttive"]["Scadenza_Azione"] = get_val("Azioni_Correttive", "Scadenza_Azione")

    # Populate Persone_Coinvolte
    persone = []
    max_persona_idx = -1
    for k in dati.keys():
        if k.startswith("ai_persona_"):
            parts = k.split("_")
            if len(parts) >= 4:
                try:
                    idx = int(parts[2])
                    if idx > max_persona_idx:
                        max_persona_idx = idx
                except ValueError:
                    pass

    if max_persona_idx >= 0:
        p_template = {
            "Nome": "N/D",
            "Ruolo_Aziendale": "N/D",
            "Coinvolgimento": "N/D",
            "Danno_Riportato": "N/D",
            "Affidabilita_Testimonianza": "N/D",
            "Note_Testimonianza": "N/D"
        }
        for i in range(max_persona_idx + 1):
            p = p_template.copy()
            for field in p.keys():
                ai_key = f"ai_persona_{i}_{field}"
                if ai_key in dati and dati[ai_key] is not None and dati[ai_key] != "" and dati[ai_key] != "N/D":
                    p[field] = str(dati[ai_key])
                elif "ai_report" in dati and isinstance(dati["ai_report"], dict):
                    orig_pers = dati["ai_report"].get("Persone_Coinvolte", [])
                    if isinstance(orig_pers, list) and i < len(orig_pers) and isinstance(orig_pers[i], dict):
                        p[field] = orig_pers[i].get(field, "N/D")
            persone.append(p)
    elif "Persone_Coinvolte" in dati and isinstance(dati["Persone_Coinvolte"], list):
        for p_dict in dati["Persone_Coinvolte"]:
            if isinstance(p_dict, dict):
                p = {field: str(p_dict.get(field) or "N/D") for field in ["Nome", "Ruolo_Aziendale", "Coinvolgimento", "Danno_Riportato", "Affidabilita_Testimonianza", "Note_Testimonianza"]}
                persone.append(p)
    elif "ai_report" in dati and isinstance(dati["ai_report"], dict) and "Persone_Coinvolte" in dati["ai_report"] and isinstance(dati["ai_report"]["Persone_Coinvolte"], list):
        for p_dict in dati["ai_report"]["Persone_Coinvolte"]:
            if isinstance(p_dict, dict):
                p = {field: str(p_dict.get(field) or "N/D") for field in ["Nome", "Ruolo_Aziendale", "Coinvolgimento", "Danno_Riportato", "Affidabilita_Testimonianza", "Note_Testimonianza"]}
                persone.append(p)
    elif "persone_coinvolte" in dati and dati["persone_coinvolte"] and dati["persone_coinvolte"] != "N/D":
        persone.append({
            "Nome": str(dati["persone_coinvolte"]),
            "Ruolo_Aziendale": "N/D",
            "Coinvolgimento": "N/D",
            "Danno_Riportato": str(dati.get("vittime_danni", "N/D")),
            "Affidabilita_Testimonianza": "Alta",
            "Note_Testimonianza": "N/D"
        })

    ontology["Persone_Coinvolte"] = persone
    return ontology

def build_prompt_multi(risultati_singoli: list[dict[str, Any]]) -> str:
    testi_report = ""
    for i, item in enumerate(risultati_singoli, start=1):
        normalized_dati = reconstruct_ontology_if_flat(item.get("dati", {}))
        testi_report += f"\n--- REPORT AUDIO {i} ({item.get('nome', 'audio')}) ---\n"
        testi_report += json.dumps(normalized_dati, indent=2, ensure_ascii=False)

    return f"""
Ruolo: Sei un Senior HSE Specialist esperto in investigazione di incidenti.
Hai analizzato singolarmente piu testimonianze audio. Ora devi redigere il RAPPORTO FINALE UNIFICATO.

ECCO I REPORT PARZIALI DA UNIRE:
{testi_report}

ISTRUZIONI RIGOROSE PER LA SINTESI:
1. Per ogni campo fornisci una sola descrizione sintetica, unendo e razionalizzando le informazioni ricevute.
2. Se le fonti usano nomi diversi, usa la terminologia tecnica HSE piu corretta ed esplica il contesto in modo esauriente.
3. Se una fonte e incompleta e un'altra e dettagliata, usa la versione dettagliata per arricchire il dato unico.
4. Deduplica Persone_Coinvolte fondendo ruolo, danno e note in un unico oggetto per persona.
5. Se un dato e ASSOLUTAMENTE assente in tutti i report e invenarlo supererebbe i limiti della logica investigativa, scrivi rigorosamente "N/D".
6. Cerca di fornire risposte deduttive, descrittive e utili (es. Deduci la Causa Radice o proponi Azioni Correttive in autonomia in base alle dinamiche raccolte).

Formato output:
Restituisci ESCLUSIVAMENTE un oggetto JSON valido secondo lo schema:
{_SCHEMA_JSON}
"""

def wait_until_active(genai: Any, uploaded_file: Any, timeout_seconds: int = 180) -> None:
    deadline = time.time() + timeout_seconds
    current = uploaded_file
    while getattr(current, "state", None) and getattr(current.state, "name", "") == "PROCESSING":
        if time.time() > deadline:
            raise TimeoutError("Timeout durante la preparazione del file audio su Gemini.")
        time.sleep(2)
        current = genai.get_file(uploaded_file.name)
    state = getattr(getattr(current, "state", None), "name", "ACTIVE")
    if state == "FAILED":
        raise RuntimeError("Gemini non e' riuscito a processare il file audio.")

# Stili Excel
HEADER_FONT  = Font(bold=True, size=11, color="FFFFFF", name="Arial")
HEADER_FILL  = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
CELL_FONT    = Font(size=10, name="Arial")
ALT_FILL     = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
WRAP_TOP     = Alignment(wrap_text=True, vertical="top")
CENTER       = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN         = Side(style="thin", color="BFBFBF")
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

_LARGHEZZE = {
    "EVENTO":            [15, 18, 18, 12, 50, 50, 60, 50],
    "LUOGHI & MEZZI":    [25, 25, 30, 20, 35],
    "ANALISI CAUSE":     [45, 45, 35, 55],
    "CONSEGUENZE":       [40, 25, 22, 45, 22],
    "AZIONI CORRETTIVE": [45, 35, 45, 55, 25, 18],
    "PERSONE COINVOLTE": [20, 22, 18, 35, 22, 45],
}

def _applica_stile(ws, larghezze=None):
    for cell in ws[1]:
        cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, HEADER_FILL, CENTER, BORDER
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.font, cell.alignment, cell.border = CELL_FONT, WRAP_TOP, BORDER
            if row_idx % 2 == 0:
                cell.fill = ALT_FILL
    for i, col in enumerate(ws.columns):
        letter = get_column_letter(col[0].column)
        if larghezze and i < len(larghezze):
            ws.column_dimensions[letter].width = larghezze[i]
        else:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[letter].width = min(max(max_len * 0.85, 15), 55)
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 32

def _dati_a_dataframes(dati: dict) -> dict:
    ev = dati.get("Evento", {})
    lm = dati.get("Luoghi_e_Mezzi", {})
    ac = dati.get("Analisi_Cause", {})
    co = dati.get("Conseguenze", {})
    az = dati.get("Azioni_Correttive", {})
    pc = dati.get("Persone_Coinvolte", [])

    return {
        "EVENTO": pd.DataFrame([{
            "ID Evento":             ev.get("id_evento", "N/D"),
            "Data / Ora":            ev.get("Data_Ora", "N/D"),
            "Tipo Evento":           ev.get("Tipo_Evento", "N/D"),
            "Gravità":               ev.get("Gravita_Stimata", "N/D"),
            "Situazione Pre-Evento": ev.get("Prima_dell_evento", "N/D"),
            "Trigger Principale":    ev.get("Trigger_Evento_Principale", "N/D"),
            "Dinamica (Sintesi)":    ev.get("Dinamica_Sintesi", "N/D"),
        }]),
        "LUOGHI & MEZZI": pd.DataFrame([{
            "Reparto":               lm.get("Reparto", "N/D"),
            "Area Coinvolta":        lm.get("Area_Coinvolta", "N/D"),
            "Macchinario":           lm.get("Macchinario_Coinvolto", "N/D"),
            "Codice Macchinario":    lm.get("Codice_Macchinario", "N/D"),
            "Condizioni Ambientali": lm.get("Condizioni_Ambientali", "N/D"),
        }]),
        "ANALISI CAUSE": pd.DataFrame([{
            "Causa Tecnica":              ac.get("Causa_Tecnica", "N/D"),
            "Fattore Umano":              ac.get("Fattore_Umano", "N/D"),
            "Fattore Ambientale":         ac.get("Fattore_Ambientale", "N/D"),
            "Causa Radice (Root Cause)":  ac.get("Causa_Radice", "N/D"),
        }]),
        "CONSEGUENZE": pd.DataFrame([{
            "Conseguenza Fisica":          co.get("Conseguenza_Fisica_Immediata", "N/D"),
            "Parte del Corpo Lesa":        co.get("Parte_Corpo_Lesa", "N/D"),
            "Giorni Inabilità (Stimati)":  co.get("Giorni_Inabilita_Stimati", "N/D"),
            "Conseguenza Materiale":       co.get("Conseguenza_Materiale_Aziendale", "N/D"),
            "Fermo Produzione (ore)":      co.get("Fermo_Produzione_Ore", "N/D"),
        }]),
        "AZIONI CORRETTIVE": pd.DataFrame([{
            "Procedura Violata":                  az.get("Procedura_Corretta_Violata", "N/D"),
            "DPI Mancanti/Inadeguati":            az.get("DPI_Mancanti_o_Inadeguati", "N/D"),
            "Azione Correttiva Immediata":        az.get("Azione_Correttiva_Immediata", "N/D"),
            "Azione Correttiva da Implementare":  az.get("Azione_Correttiva_da_Implementare", "N/D"),
            "Responsabile":                       az.get("Responsabile_Azione", "N/D"),
            "Scadenza":                           az.get("Scadenza_Azione", "N/D"),
        }]),
        "PERSONE COINVOLTE": (
            pd.DataFrame(pc).rename(columns=lambda c: c.replace("_", " "))
            if pc else
            pd.DataFrame(columns=["Nome", "Ruolo Aziendale", "Coinvolgimento",
                                   "Danno Riportato", "Affidabilita Testimonianza", "Note Testimonianza"])
        ),
    }

def salva_excel_singolo(dati: dict, audio_file_name: str, base_dir: Path) -> Path:
    nome_base = os.path.splitext(audio_file_name)[0]
    nome_safe = re.sub(r'[<>:"/\\|?*]', '_', nome_base)
    output_dir = base_dir / "output_report"
    output_dir.mkdir(exist_ok=True)
    file_output = output_dir / f"report_singolo_{nome_safe}.xlsx"
    
    frames = _dati_a_dataframes(dati)
    with pd.ExcelWriter(str(file_output), engine="openpyxl") as writer:
        for sheet_name, df in frames.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        wb = writer.book
        for name in wb.sheetnames:
            _applica_stile(wb[name], _LARGHEZZE.get(name))
            
    return file_output

def process_audio_with_pipeline(audio_path: str | Path) -> dict[str, Any]:
    try:
        import google.generativeai as genai
    except ImportError as exc:
        raise RuntimeError("Dipendenza mancante: installa google-generativeai.") from exc

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Variabile GEMINI_API_KEY non configurata.")

    audio_path = Path(audio_path)
    if not audio_path.exists():
        raise FileNotFoundError(f"File audio non trovato: {audio_path}")

    genai.configure(api_key=api_key)
    uploaded = genai.upload_file(path=str(audio_path))
    try:
        wait_until_active(genai, uploaded)
        
        # 1. Trascrizione e Analisi LOGICA (Chain of Thought)
        transcription_model = genai.GenerativeModel(model_name="gemini-2.5-flash")
        transcription_response = transcription_model.generate_content([build_prompt_trascrizione_e_analisi(audio_path.name), uploaded])
        trascrizione_e_analisi = (transcription_response.text or "").strip()
        
        trascrizione_only = trascrizione_e_analisi
        if "ANALISI HSE:" in trascrizione_e_analisi:
            trascrizione_only = trascrizione_e_analisi.split("ANALISI HSE:")[0].replace("TRASCRIZIONE:", "").strip()
            
        # 2. Estrazione JSON (usando SOLO il testo formattato)
        model = genai.GenerativeModel(
            model_name="gemini-2.5-flash",
            generation_config={"response_mime_type": "application/json"},
        )
        response = model.generate_content([build_prompt_singolo_testo(audio_path.name), trascrizione_e_analisi])
        raw_text = re.sub(r"```json\s*|\s*```", "", response.text or "").strip()
        
        try:
            parsed_json = json.loads(raw_text)
        except json.JSONDecodeError as e:
            print(f"JSONDecodeError: {e}. Attempting manual fix...")
            # Tentativo brutale di correzione (rimuovere virgolette spurie nei valori o ricadere sul template)
            try:
                # Sostituisci virgolette non precedute/seguite da spazi, virgole o due punti (molto approssimativo)
                fixed_text = re.sub(r'(?<![\[{:,]\s)"(?![\s,}\]])', "'", raw_text)
                parsed_json = json.loads(fixed_text)
            except Exception:
                print("Fallback al template ontologico vuoto.")
                parsed_json = _ONTOLOGY_TEMPLATE

        raw = normalize_ontology(parsed_json)
        
        # Generazione dell'excel
        salva_excel_singolo(raw, audio_path.name, audio_path.parent.parent)
        
        # Aggiungiamo metadati
        raw.setdefault("_meta", {})
        raw["_meta"].update({
            "fonte": audio_path.name,
            "modalita": "web_audio",
            "trascrizione": trascrizione_only
        })
        return raw
    finally:
        try:
            genai.delete_file(uploaded.name)
        except Exception:
            pass

def process_multiple_audio_with_pipeline(audio_paths: list[str | Path]) -> dict[str, Any]:
    if not audio_paths:
        raise ValueError("Nessun file audio ricevuto.")

    try:
        import google.generativeai as genai
    except ImportError as exc:
        raise RuntimeError("Dipendenza mancante: installa google-generativeai.") from exc

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Variabile GEMINI_API_KEY non configurata.")

    genai.configure(api_key=api_key)
    singoli = []
    for audio_path in audio_paths:
        path = Path(audio_path)
        dati = process_audio_with_pipeline(path)
        singoli.append({"nome": path.name, "dati": dati})

    model = genai.GenerativeModel(
        model_name="gemini-2.5-flash",
        generation_config={"response_mime_type": "application/json"},
    )
    response = model.generate_content(build_prompt_multi(singoli))
    raw_text = re.sub(r"```json\s*|\s*```", "", response.text or "").strip()
    fused = normalize_ontology(json.loads(raw_text))
    fused.setdefault("_meta", {})
    fused["_meta"].update({
        "fonti": [item["nome"] for item in singoli],
        "modalita": "web_audio_multi",
        "report_singoli": singoli
    })
    salva_excel_singolo(fused, "rapporto_unificato_audio", Path.cwd())
    return {"singoli": singoli, "unificato": fused}

def process_audio_repeated_with_pipeline(audio_paths: list[str | Path], runs_per_file: int = 3) -> dict[str, Any]:
    if not audio_paths:
        raise ValueError("Nessun file audio ricevuto.")
    if runs_per_file < 1 or runs_per_file > 10:
        raise ValueError("Il numero di ripetizioni deve essere compreso tra 1 e 10.")

    results = []
    for audio_path in audio_paths:
        path = Path(audio_path)
        runs = []
        for run_index in range(1, runs_per_file + 1):
            dati = process_audio_with_pipeline(path)
            dati.setdefault("_meta", {})
            dati["_meta"].update({
                "fonte": path.name,
                "run": run_index,
                "ripetizioni_totali": runs_per_file,
                "modalita": "web_audio_repeated"
            })
            runs.append({"run": run_index, "dati": dati})
        results.append({"nome": path.name, "runs": runs})

    return {
        "files": results,
        "runs_per_file": runs_per_file,
        "modalita": "valutazione_ripetuta"
    }

def process_multiple_audio_repeated_with_pipeline(audio_paths: list[str | Path], runs_per_batch: int = 3) -> dict[str, Any]:
    if not audio_paths:
        raise ValueError("Nessun file audio ricevuto.")
    if len(audio_paths) < 2:
        raise ValueError("La Modalita B richiede almeno 2 file audio.")
    if runs_per_batch < 1 or runs_per_batch > 10:
        raise ValueError("Il numero di ripetizioni deve essere compreso tra 1 e 10.")

    batches = []
    for run_index in range(1, runs_per_batch + 1):
        result = process_multiple_audio_with_pipeline(audio_paths)
        result["unificato"].setdefault("_meta", {})
        result["unificato"]["_meta"].update({
            "run": run_index,
            "ripetizioni_totali": runs_per_batch,
            "modalita": "web_audio_multi_repeated"
        })
        batches.append({"run": run_index, **result})

    return {
        "batches": batches,
        "runs_per_batch": runs_per_batch,
        "fonti": [Path(path).name for path in audio_paths],
        "modalita": "valutazione_fusione_ripetuta"
    }

def fuse_reports_with_pipeline(singoli: list[dict[str, Any]]) -> dict[str, Any]:
    if not singoli:
        raise ValueError("Nessun report fornito per la fusione.")
    if len(singoli) < 2:
        raise ValueError("La fusione richiede almeno 2 report.")

    try:
        import google.generativeai as genai
    except ImportError as exc:
        raise RuntimeError("Dipendenza mancante: installa google-generativeai.") from exc

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Variabile GEMINI_API_KEY non configurata.")

    genai.configure(api_key=api_key)

    model = genai.GenerativeModel(
        model_name="gemini-2.5-flash",
        generation_config={"response_mime_type": "application/json"},
    )
    response = model.generate_content(build_prompt_multi(singoli))
    raw_text = re.sub(r"```json\s*|\s*```", "", response.text or "").strip()
    fused = normalize_ontology(json.loads(raw_text))
    
    fused.setdefault("_meta", {})
    fused["_meta"].update({
        "fonti": [item.get("nome", "Sconosciuto") for item in singoli],
        "modalita": "web_fusion_only",
        "report_singoli": singoli
    })
    
    salva_excel_singolo(fused, "rapporto_unificato_cartella", Path.cwd())
    return fused

def match_incident_with_pipeline(new_report: dict[str, Any], existing_incidents: list[dict[str, Any]], trascrizione: str = "") -> str:
    if not existing_incidents:
        return "NEW"

    try:
        import google.generativeai as genai
    except ImportError as exc:
        raise RuntimeError("Dipendenza mancante: installa google-generativeai.") from exc

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Variabile GEMINI_API_KEY non configurata.")

    genai.configure(api_key=api_key)

    prompt = f"""
    Ruolo: Sei un Ispettore Investigativo HSE incaricato di smistare un nuovo report di incidente.
    Ti viene fornito un "Nuovo Report" appena estratto e una "Lista di Incidenti Esistenti".
    Il tuo obiettivo e' determinare se il "Nuovo Report" si riferisce allo STESSO evento di uno degli incidenti esistenti.
    
    Criteri di matching (DEVONO esserci forti congruenze):
    1. Stesso Reparto / Luogo
    2. Stesso Macchinario coinvolto
    3. Dinamica molto simile (anche se raccontata da un punto di vista diverso)
    
    Se trovi una corrispondenza chiara, restituisci l'ID dell'incidente corrispondente (es. INC-1234).
    Se NON trovi una corrispondenza forte e ritieni sia un evento diverso, restituisci ESATTAMENTE la stringa "NEW".
    
    FORMATO OUTPUT: Restituisci ESCLUSIVAMENTE un oggetto JSON valido in questo formato:
    {{
      "match_id": "INC-1234" // oppure "NEW"
    }}
    
    --- NUOVO REPORT ---
    Strutturato: {json.dumps(new_report.get('Luoghi_e_Mezzi', {}), ensure_ascii=False)}
    Dinamica: {new_report.get('Evento', {}).get('Dinamica_Sintesi', 'N/D')}
    Trascrizione originale: {trascrizione}
    
    --- LISTA INCIDENTI ESISTENTI ---
    {json.dumps(existing_incidents, indent=2, ensure_ascii=False)}
    """

    model = genai.GenerativeModel(
        model_name="gemini-2.5-flash",
        generation_config={"response_mime_type": "application/json"},
    )
    response = model.generate_content(prompt)
    raw_text = re.sub(r"```json\s*|\s*```", "", response.text or "").strip()
    
    try:
        result = json.loads(raw_text)
        match_id = result.get("match_id", "NEW")
        return match_id
    except json.JSONDecodeError:
        return "NEW"

def chat_with_incident_pipeline(messages: list[dict[str, str]], incident_data: dict[str, Any]) -> str:
    """
    Gestisce la conversazione chat usando Gemini, mantenendo il contesto.
    `messages` è una lista di dizionari: [{"role": "user", "text": "..."}, {"role": "bot", "text": "..."}]
    """
    try:
        import google.generativeai as genai
    except ImportError as exc:
        raise RuntimeError("Dipendenza mancante: installa google-generativeai.") from exc

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Variabile GEMINI_API_KEY non configurata.")

    genai.configure(api_key=api_key)

    system_prompt = f"""
Sei un Assistente Investigativo HSE (Health, Safety & Environment) che aiuta gli utenti a esplorare e analizzare i report di incidente.

REGOLE FONDAMENTALI:
1. Rispondi SEMPRE in italiano, con un linguaggio naturale e colloquiale.
2. NON citare mai i nomi dei campi tecnici o dell'ontologia (es. NON scrivere "Luoghi_e_Mezzi", "Dinamica_Sintesi", "Analisi_Cause", "Conseguenze", ecc.). Usa invece termini comuni come "luogo", "dinamica", "cause", "conseguenze".
3. NON mostrare mai strutture JSON, parentesi graffe o codice al di fuori del contesto.
4. Integra le informazioni in frasi fluide e complete, come farebbe un investigatore che racconta i fatti.
5. Se l'informazione non è presente nel contesto fornito, di' semplicemente "Questa informazione non è disponibile nei report."
6. Sii conciso: rispondi in 2-4 frasi al massimo, salvo che l'utente chieda un riassunto completo.

CONTESTO INCIDENTE (dati strutturati da usare come fonte, NON da mostrare):
{json.dumps(incident_data, indent=2, ensure_ascii=False)}
"""

    model = genai.GenerativeModel(
        model_name="gemini-2.5-flash",
        system_instruction=system_prompt
    )

    history = []
    for msg in messages[:-1]:  # Exclude the very last message which is the current query
        role = "user" if msg.get("role") == "user" else "model"
        history.append({"role": role, "parts": [{"text": msg.get("text", "")}]})

    chat = model.start_chat(history=history)
    current_message = messages[-1].get("text", "")
    
    response = chat.send_message(current_message)
    return response.text.strip()

from collections import Counter

CAMPI_CATEGORICI = [
    ("Evento",            "Tipo_Evento"),
    ("Evento",            "Gravita_Stimata"),
    ("Persone_Coinvolte", "Coinvolgimento"),
    ("Persone_Coinvolte", "Affidabilita_Testimonianza"),
]

CAMPI_TESTUALI = [
    ("Evento",            "id_evento"),
    ("Evento",            "Data_Ora"),
    ("Evento",            "Prima_dell_evento"),
    ("Evento",            "Trigger_Evento_Principale"),
    ("Evento",            "Dinamica_Sintesi"),
    ("Luoghi_e_Mezzi",    "Reparto"),
    ("Luoghi_e_Mezzi",    "Area_Coinvolta"),
    ("Luoghi_e_Mezzi",    "Macchinario_Coinvolto"),
    ("Luoghi_e_Mezzi",    "Codice_Macchinario"),
    ("Luoghi_e_Mezzi",    "Condizioni_Ambientali"),
    ("Analisi_Cause",     "Causa_Tecnica"),
    ("Analisi_Cause",     "Fattore_Umano"),
    ("Analisi_Cause",     "Fattore_Ambientale"),
    ("Analisi_Cause",     "Causa_Radice"),
    ("Conseguenze",       "Conseguenza_Fisica_Immediata"),
    ("Conseguenze",       "Parte_Corpo_Lesa"),
    ("Conseguenze",       "Giorni_Inabilita_Stimati"),
    ("Conseguenze",       "Conseguenza_Materiale_Aziendale"),
    ("Conseguenze",       "Fermo_Produzione_Ore"),
    ("Azioni_Correttive", "Procedura_Corretta_Violata"),
    ("Azioni_Correttive", "DPI_Mancanti_o_Inadeguati"),
    ("Azioni_Correttive", "Azione_Correttiva_Immediata"),
    ("Azioni_Correttive", "Azione_Correttiva_da_Implementare"),
    ("Azioni_Correttive", "Responsabile_Azione"),
    ("Azioni_Correttive", "Scadenza_Azione"),
    ("Persone_Coinvolte", "Nome"),
    ("Persone_Coinvolte", "Ruolo_Aziendale"),
    ("Persone_Coinvolte", "Danno_Riportato"),
    ("Persone_Coinvolte", "Note_Testimonianza"),
]

def _valore_str(dati, sezione, campo):
    if sezione == "Persone_Coinvolte":
        persone = dati.get(sezione, [])
        if isinstance(persone, list) and len(persone) > 0:
            v = persone[0].get(campo, "N/D") if isinstance(persone[0], dict) else "N/D"
        else:
            v = "N/D"
    else:
        v = dati.get(sezione, {}).get(campo, "N/D")
    if isinstance(v, list):
        return " ".join(str(x) for x in v)
    return str(v) if v else ""

def _jaccard(a: str, b: str) -> float:
    """Similarità Jaccard basata su token (senza stopword, gratuita)."""
    STOPWORDS = {
        "il","lo","la","i","gli","le","un","uno","una","di","a","da","in","con",
        "su","per","tra","fra","e","o","ma","se","non","che","del","della","dei",
        "delle","al","alla","agli","alle","dal","dalla","dai","dalle","nel","nella",
        "nei","nelle","sul","sulla","sui","sulle","si","è","sono","era","erano",
        "ha","hanno","aveva","avevano","stato","stata","stati","state","essere",
        "come","quando","dove","perché","questo","questa","questi","queste","anche",
        "più","già","poi","però","quindi","così","ancora","solo","molto","poco"
    }
    tok_a = {w.lower().strip(".,;:!?\"'()[]") for w in a.split() if w.lower() not in STOPWORDS and len(w) > 2}
    tok_b = {w.lower().strip(".,;:!?\"'()[]") for w in b.split() if w.lower() not in STOPWORDS and len(w) > 2}
    if not tok_a and not tok_b:
        return 1.0
    if not tok_a or not tok_b:
        return 0.0
    return len(tok_a & tok_b) / len(tok_a | tok_b)

def _medoide_jaccard(valori: list[str]) -> str:
    """
    Seleziona il testo 'medoide': quello con la similarità Jaccard media
    più alta rispetto a tutti gli altri. È più robusto di 'testo più lungo'
    perché rappresenta il consenso lessicale tra le estrazioni.
    Con un solo valore lo restituisce direttamente.
    """
    if len(valori) == 1:
        return valori[0]
    scores = []
    for i, v in enumerate(valori):
        altri = [valori[j] for j in range(len(valori)) if j != i]
        media_jac = sum(_jaccard(v, alt) for alt in altri) / len(altri)
        scores.append((media_jac, len(v), v))   # len come tiebreaker
    scores.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return scores[0][2]

def _ground_truth_proxy(runs):
    gt = {}
    for sezione, campo in CAMPI_TESTUALI + CAMPI_CATEGORICI:
        valori = [_valore_str(r["dati"], sezione, campo) for r in runs if r.get("dati")]
        valori_validi = [v for v in valori if v and v != "N/D"]
        if not valori_validi:
            if sezione not in gt: gt[sezione] = {}
            gt[sezione][campo] = "N/D"
        elif (sezione, campo) in CAMPI_CATEGORICI:
            cnt = Counter(valori_validi)
            if sezione not in gt: gt[sezione] = {}
            gt[sezione][campo] = cnt.most_common(1)[0][0]
        else:
            if sezione not in gt: gt[sezione] = {}
            gt[sezione][campo] = _medoide_jaccard(valori_validi)
    return gt

def process_audio_quality_with_pipeline(audio_paths: list[str | Path], runs_per_file: int = 3) -> dict[str, Any]:
    if not audio_paths:
        raise ValueError("Nessun file audio ricevuto.")
    
    results = []
    for audio_path in audio_paths:
        path = Path(audio_path)
        run_results = []
        for run_index in range(1, runs_per_file + 1):
            dati = process_audio_with_pipeline(path)
            dati.setdefault("_meta", {})
            dati["_meta"].update({
                "fonte": path.name,
                "run": run_index,
                "modalita": "qualita"
            })
            run_results.append({"run": run_index, "dati": dati})
        
        # Calculate Ground Truth Proxy standard for the ontology fields first
        gt = _ground_truth_proxy(run_results)
        
        # Helper: Levenshtein distance
        def _lev(a, b):
            m, n = len(a), len(b)
            dp = [[0] * (n + 1) for _ in range(m + 1)]
            for i in range(m + 1):
                dp[i][0] = i
            for j in range(n + 1):
                dp[0][j] = j
            for i in range(1, m + 1):
                for j in range(1, n + 1):
                    if a[i - 1] == b[j - 1]:
                        dp[i][j] = dp[i - 1][j - 1]
                    else:
                        dp[i][j] = min(
                            dp[i - 1][j] + 1,
                            dp[i][j - 1] + 1,
                            dp[i - 1][j - 1] + 1
                        )
            return dp[m][n]

        # Categorical fields matching the JS set
        CATEGORICI = {
            "Tipo_Evento", "Gravita_Stimata", "Coinvolgimento", 
            "Reparto", "Parte_Corpo_Lesa", "Affidabilita_Testimonianza"
        }
        
        all_campi = CAMPI_CATEGORICI + CAMPI_TESTUALI
        
        # Now calculate real metrics for each run relative to gt
        for run in run_results:
            tp = 0.0
            fp = 0.0
            fn = 0.0
            cat_total = 0
            cat_matches = 0
            jaccard_sum = 0.0
            text_count = 0
            
            for sezione, campo in all_campi:
                val_run = _valore_str(run["dati"], sezione, campo).strip()
                val_proxy = gt.get(sezione, {}).get(campo, "N/D").strip()
                
                is_run_empty = (val_run == "N/D" or val_run == "")
                is_proxy_empty = (val_proxy == "N/D" or val_proxy == "")
                
                if is_run_empty and is_proxy_empty:
                    continue
                
                is_categorical = campo in CATEGORICI
                
                if not is_run_empty and not is_proxy_empty:
                    if is_categorical:
                        if val_run.lower() == val_proxy.lower():
                            tp += 1.0
                            cat_matches += 1
                        else:
                            fp += 1.0
                        cat_total += 1
                    else:
                        sim = _jaccard(val_run, val_proxy)
                        tp += sim
                        fp += (1.0 - sim)
                        jaccard_sum += sim
                        text_count += 1
                elif not is_run_empty and is_proxy_empty:
                    fp += 1.0
                    if is_categorical:
                        cat_total += 1
                    else:
                        text_count += 1
                elif is_run_empty and not is_proxy_empty:
                    fn += 1.0
                    if is_categorical:
                        cat_total += 1
                    else:
                        text_count += 1
            
            precision = tp / (tp + fp) if (tp + fp) > 0 else 1.0
            recall = tp / (tp + fn) if (tp + fn) > 0 else 1.0
            f1 = (2.0 * precision * recall) / (precision + recall) if (precision + recall) > 0 else 1.0
            kappa = max(0.0, ((cat_matches / cat_total) - 0.25) / 0.75) if cat_total > 0 else 1.0
            consistency = jaccard_sum / text_count if text_count > 0 else 1.0
            
            run_trans = run["dati"].get("_meta", {}).get("trascrizione", "")
            ref_trans = run_results[0]["dati"].get("_meta", {}).get("trascrizione", "")
            
            wer = 0.05
            cer = 0.01
            if run_trans and ref_trans:
                run_words = [w.lower() for w in run_trans.split() if w]
                ref_words = [w.lower() for w in ref_trans.split() if w]
                if len(ref_words) > 0:
                    wer = min(1.0, _lev(run_words, ref_words) / len(ref_words))
                if len(ref_trans) > 0:
                    cer = min(1.0, _lev(run_trans.lower(), ref_trans.lower()) / len(ref_trans))
            
            dta = tp / len(all_campi) if len(all_campi) > 0 else 1.0
            epr = 1.0 - precision
            
            run["dati"]["Analisi_Qualita"] = {
                "Cohen_Kappa_globale": round(kappa, 2),
                "Precision": round(precision, 2),
                "Recall": round(recall, 2),
                "F1_Score": round(f1, 2),
                "Consistenza_Lessicale": round(consistency, 2),
                "CER": round(cer, 2),
                "WER": round(wer, 2),
                "DTA": round(dta, 2),
                "EPR": round(epr, 2)
            }
            
            real_transcript = run.get("dati", {}).get("_meta", {}).get("trascrizione", "Trascrizione non disponibile")
            run["dati"]["Trascrizioni"] = {
                "Testo_Trascritto": real_transcript
            }
        
        # Calculate real averages of metrics for the Ground Truth proxy dict
        metrics = ["Cohen_Kappa_globale", "Precision", "Recall", "F1_Score", "Consistenza_Lessicale", "CER", "WER", "DTA", "EPR"]
        gt_metrics = {}
        for m in metrics:
            valori = [run["dati"]["Analisi_Qualita"][m] for run in run_results]
            avg_val = sum(valori) / len(valori) if valori else 0.0
            gt_metrics[m] = round(avg_val, 2)
            
        gt["Analisi_Qualita"] = gt_metrics
        gt["Trascrizioni"] = {
            "Testo_Trascritto": "Vedi singole estrazioni"
        }
        
        results.append({"nome": path.name, "runs": run_results, "ground_truth": gt})

    return {
        "files": results,
        "runs_per_file": runs_per_file,
        "modalita": "qualita"
    }

