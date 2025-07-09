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
oppure 
```
.\venv\Scripts\activate
```

Installaci dentro le librerie che servono:
```shell
pip install -r requirements.txt
```

modifica i parametri dello script e poi lancialo con:
```shell
python3 job
```

## job2: scheda riepilogo

Dato un file excel e un template (`input_file`, `template_file`) nella cartella principale, genera una copia del `template_file` per ogni riga di `input_file`, rimpiazzando i valori definiti in `replace_fields`.

Il nome file è di ogni copia è `sede_nome_cognome.docx`.
