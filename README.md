## ASC jobs

Crea l'env:

```shell
python3 -m venv venv
```

Poi attivalo con uno di questi comandi:
```shell
. venv/bin/activate
source venv/bin/activate
venv\Scripts\activate
```

Installaci dentro le librerie che servono:
```shell
python3 -m pip install pandas, openpyxl, python-docx 
```

Carica i file dentro la cartella files,
modifica i parametri dentro il job, 
e poi lancia il job
```shell
python3 job2
```

## job2: scheda riepilogo

Dato un file excel e un template (`input_file`, `template_file`) nella cartella `file`, genera una copia del `template_file` nella cartella `file`, per ogni riga di `input_file` riempiendo i valori definiti in `replace_fields` con quelli trovati nella riga di `input_file`.

Il nome file è di ogni copia è `sede_nome_cognome.docx`.
