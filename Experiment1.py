import os
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# Funktion zur Berechnung der Kosinus-Ähnlichkeit
def calculate_cosine_similarity(text1, text2):
    vectorizer = TfidfVectorizer()
    tfidf_matrix = vectorizer.fit_transform([text1, text2])
    cosine_sim = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])
    return cosine_sim[0][0]

def evaluate_answer(model_answer, goldstandards, threshold=0.6):
    max_similarity = 0
    for gold in goldstandards:
        similarity = calculate_cosine_similarity(model_answer, gold)
        max_similarity = max(max_similarity, similarity)
    return 1 if max_similarity >= threshold else "N/A"

def categorize_error(model_answer, goldstandard, threshold_lexical=0.8, threshold_semantic=0.6):
    similarity = calculate_cosine_similarity(model_answer, goldstandard)
    if similarity >= threshold_lexical:
        return 'Lexikalisch'
    elif similarity >= threshold_semantic:
        return 'Semantisch'
    else:
        return 'Syntaktisch'

# Lade die Daten aus einer Excel-Datei
data = pd.read_excel("C:/Users/melis/OneDrive/Dokumente/Uni final/Hausarbeit/Auswertung_Ambiguität.xlsx")

# Stelle sicher, dass keine leeren Zellen vorhanden sind
data = data.fillna("N/A")

# Berechne die Beurteilungen und Fehlerkategorisierungen für jedes Modell
for i in range(1, 3):  # Nur die ersten beiden Antwortspalten berücksichtigen (wenn du mehr hast, anpassen)
    # Beurteilungen
    data['ChatGPT-Beurteilung {}'.format(i)] = data.apply(lambda row, i=i: evaluate_answer(row['ChatGPT-Ausgabe {}'.format(i)], 
                                            [row['Goldstandard-Interpretation {}'.format(i)], 
                                             row['Goldstandard-Interpretation 2'], 
                                             row['Goldstandard-Interpretation 3']]), axis=1)
    data['Gemini-Beurteilung {}'.format(i)] = data.apply(lambda row, i=i: evaluate_answer(row['Gemini-Ausgabe {}'.format(i)], 
                                            [row['Goldstandard-Interpretation {}'.format(i)], 
                                             row['Goldstandard-Interpretation 2'], 
                                             row['Goldstandard-Interpretation 3']]), axis=1)
    data['Microsoft Copilot-Beurteilung {}'.format(i)] = data.apply(lambda row, i=i: evaluate_answer(row['Microsoft Copilot-Ausgabe {}'.format(i)], 
                                            [row['Goldstandard-Interpretation {}'.format(i)], 
                                             row['Goldstandard-Interpretation 2'], 
                                             row['Goldstandard-Interpretation 3']]), axis=1)

    # Fehlerkategorisierungen
    data[f'ChatGPT-Fehlerkategorie {i}'] = data.apply(lambda row, i=i: categorize_error(row[f'ChatGPT-Ausgabe {i}'], row[f'Goldstandard-Interpretation {i}']), axis=1)
    data[f'Gemini-Fehlerkategorie {i}'] = data.apply(lambda row, i=i: categorize_error(row[f'Gemini-Ausgabe {i}'], row[f'Goldstandard-Interpretation {i}']), axis=1)
    data[f'Microsoft Copilot-Fehlerkategorie {i}'] = data.apply(lambda row, i=i: categorize_error(row[f'Microsoft Copilot-Ausgabe {i}'], row[f'Goldstandard-Interpretation {i}']), axis=1)

# Speichern der Datei an einem bekannten Ort
output_path = "C:/Users/melis/OneDrive/Dokumente/ausgewertete_tabelle.xlsx"  # Pfad anpassen
try:
    data.to_excel(output_path, index=False)
    print(f"Die Auswertung wurde abgeschlossen und in '{output_path}' gespeichert.")
except Exception as e:
    print(f"Fehler beim Speichern der Excel-Datei: {e}")

import os
output_path = "ausgewertete_tabelle.xlsx"  # Datei im gleichen Ordner wie das Skript speichern
