# Klikátko

Klikátko připraví Word s osnovou průvodního listu a souhrnné technické zprávy. Ty pak jen dopíšeš odborný obsah.

## Jak to spustit

1. Dvakrát klikni na `spustit.bat`.
2. Poprvé to chvíli trvá. Počítač si sám stáhne Python a potřebné věci. Potřebuje k tomu internet, ale jen poprvé. Může se objevit okno s dotazem na povolení. Klikni na Ano.
3. Otevře se prohlížeč s Klikátkem.
4. Černé okno nech otevřené, dokud pracuješ. Až skončíš, zavři ho křížkem.

## Jak na to

Nahoře vidíš silnici se třemi zastávkami. Auto ukazuje, kde zrovna jsi. Na zastávku můžeš kdykoli kliknout.

1. **Údaje o stavbě.** Vyplň název stavby, číslo zakázky a další údaje. Vpravo hned vidíš, jak bude vypadat titulní strana. Obrázek na titulní straně vyměníš tlačítkem **Změnit obrázek**.
2. **Kapitoly.** Zaškrtni kapitoly, které do zprávy patří. Šipkou rozbalíš podkapitoly. Když klikneš na název kapitoly, uvidíš, co se do ní píše.
3. **Word.** Vyber složku zakázky a klikni na **Vytvořit Word**. Pak můžeš kliknout na **Otevřít ve Wordu** nebo **Otevřít složku**.

Některé věci se hodí vědět:

- Tlačítko **Uložit** nahoře uloží rozdělanou zakázku. Najdeš ji pak v **Uložených zakázkách**.
- Když často vybíráš stejné kapitoly, klikni u kapitol na **Uložit tento výběr**. Příště ho načteš tlačítkem **Načíst předvolbu**.
- Klikátko nikdy nepřepíše soubor, který už existuje. Nová verze dostane do názvu `_02`, další `_03` a tak dál.

Ruční úpravy hotového souboru ve Wordu už Klikátko nenačte zpátky.

## Kde jsou moje data

Zakázky, předvolby, obrázky titulní strany i hotové dokumenty leží ve složce `C:\Users\<jméno>\Klikatko-data`. Hotové dokumenty jsou v její podsložce `reports`.

## Když něco nejde

- **Černé okno se hned zavře, nebo se nic neotevře.** Podívej se do černého okna, co tam stojí. Zavolej Barboře a řekni jí to.
- **Prohlížeč se sám neotevřel.** Otevři ho ručně a do adresního řádku napiš `http://127.0.0.1:8765`. Stejnou adresu najdeš v černém okně.
- **Stránka hlásí „Obnovte stránku“.** Zmáčkni F5.
- **Při prvním spuštění se nic neděje.** Počkej pár minut. Instalace Pythonu chvíli trvá. Pokud to dál nejde, zavolej Barboře.

---

## Pro Barboru (technické)

Klikátko je lokální aplikace. Spustí malý HTTP server jen na počítači, kde běží, a otevře ho ve výchozím prohlížeči. Nevyžaduje účet. Po instalaci závislostí funguje offline.

Požadavky:

- Python 3.10 nebo novější. `spustit.bat` ho při prvním spuštění nainstaluje přes `winget` (Python 3.12, pro aktuálního uživatele). Bez `winget` nebo bez internetu skript skončí hlásením, že má volat Barboru.
- Závislosti z `requirements.txt`: `python-docx` a `Pillow`.

Ruční spuštění v kořenové složce projektu:

```text
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

V Linuxu nebo macOS použij `.venv/bin/python`.

Volby:

```text
python main.py --no-browser
python main.py --port 8766
python main.py --data-dir "C:\Cesta\k\datům"
```

Volby lze kombinovat. Aplikaci ukončíš v okně terminálu klávesou `Ctrl+C`.

Data: výchozí umístění je `~/Klikatko-data`. Změníš ho volbou `--data-dir` nebo proměnnou prostředí `KLIKATKO_DATA_DIR`.

Prohlížeč a server musí běžet na stejném počítači. Cesta k souboru zadaná v aplikaci označuje soubor na počítači se serverem. V cloudovém prostředí server složky na notebooku neuvidí. Stažené dokumenty pak přenes na notebook ručně.

Osnova v `klikatko/catalog.json` vychází z dodaného vzoru. Její právní úplnost nebyla ověřena. Před použitím dokumentace zkontroluj, že odpovídá požadavkům konkrétního záměru.

Aplikace zachovává prvky titulní strany předlohy, včetně jejího písma a ukotveného obrázku. Po vytvoření dokument otevři ve Wordu a zkontroluj sazbu, zalomení a titulní stranu.

Písma Atkinson Hyperlegible Next a Osifont jsou přibalená ve složce `klikatko/web/static/fonts`, takže aplikace vypadá stejně i bez internetu.
