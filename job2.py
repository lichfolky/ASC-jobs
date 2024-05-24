import pandas as pd
from docx import Document

dir = "files/"
output_dir = "output/"
input_file = "VOLONTAR3 2024-25.xlsx"
template_file = "scheda_riepilogativa_struttura_ASC.docx"

## testo sul doc : colonna sul excel
replace_fields = {
    "Data Avvio:": "Data Avvio:",
    "Titolo del progetto:": "Titolo del progetto:",
    "Sede di attuazione:": "Sede di attuazione:",
    "Operatore Locale di Progetto:": "Operatore Locale di Progetto:",
    "Telefono olp:": "Telefono olp:",
    "e.mail olp:": "e.mail olp:",
    "Docenti di formazione specifica (ci potranno essere delle variazioni in seguito a cambi di formatorɜ in corso di progetto):": "Docente/i di formazione specifica:",
}

check = []


## sostituisce nel file new_document una word con replacement
def replace(new_document, word, replacement):
    # print("replace", word, replacement)
    for p in new_document.paragraphs:
        if p.text.find(word) >= 0:
            p.text = p.text.replace(word, replacement)

    for table in new_document.tables:
        for r in table.rows:
            for c in r.cells:
                if c.text.find(word) >= 0:
                    c.text = c.text.replace(word, replacement)


# Crea il file "sede_nome_cognome.docx" riempiendo gli spazi vuoti
def create_file(row):
    nome = row.iloc[1]
    cognome = row.iloc[2]
    sede = str(row.loc["Sede di attuazione:"]).replace("/", "-")
    filename = sede + "_" + nome + "_" + cognome + ".docx"

    new_document = Document(dir + template_file)
    replace(
        new_document,
        "Nominativo op. vol.:",
        "Nominativo op. vol.: " + nome + " " + cognome,
    )
    for word in replace_fields:

        value = str(row.loc[replace_fields[word]]).strip()
        if value == "" or value == "nan":
            check.append(nome + " " + cognome + " manca " + replace_fields[word])
        else:
            replace(new_document, word, word + " " + value)

    new_document.save(output_dir + filename)


## main

df = pd.read_excel(dir + input_file, 1)
# df.head(10).apply(create_file, axis=1)
df.apply(create_file, axis=1)

str = "\n"
print(str.join(check))
