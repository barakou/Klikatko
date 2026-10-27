# Klikátko

Klikátko je lokální aplikace pro přípravu osnovy a dokumentu Word se souhrnnou technickou zprávou. Spustí malý HTTP server pouze na počítači, kde běží, a otevře aplikaci v jeho výchozím prohlížeči. Po instalaci závislostí aplikace funguje offline a nevyžaduje účet.

## Požadavky

- Python 3.10 nebo novější
- Závislosti z `requirements.txt`: `python-docx` a `Pillow`

## Spuštění ve Windows

Dvakrát klikněte na `spustit.bat`. Při prvním spuštění skript vytvoří virtuální prostředí `.venv` a stáhne potřebné balíčky (je potřeba internetové připojení). Po jejich instalaci aplikace funguje offline. Další spuštění použijí toto prostředí a aplikaci rovnou spustí.

## Spuštění ručně

V kořenové složce projektu spusťte:

```text
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

V Linuxu nebo macOS nahraďte poslední cestu k Pythonu výrazem `.venv/bin/python`.

Aplikace ve výchozím nastavení sama otevře prohlížeč na místní adrese. Dostupné volby:

```text
python main.py --no-browser
python main.py --port 8766
python main.py --data-dir "C:\Cesta\k\datům"
```

Volby lze kombinovat. Aplikaci ukončíte v okně terminálu stisknutím `Ctrl+C`.

## Data a výstupy

Zakázky, předvolby, obrázky titulní strany a vytvořené dokumenty se ukládají do složky `~/Klikatko-data`; na Windows jde obvykle o `C:\Users\<uživatel>\Klikatko-data`. Dokumenty se standardně ukládají do její podsložky `reports`. Umístění dat lze změnit volbou `--data-dir` nebo proměnnou prostředí `KLIKATKO_DATA_DIR`.

Prohlížeč a server musí běžet na stejném počítači. Cesta k souboru zadaná v aplikaci označuje soubor na počítači, kde běží server. Pokud aplikaci spustíte v cloudovém prostředí, server neuvidí složky na vašem notebooku; stažené dokumenty pak přeneste do požadované složky na notebooku.

Každé nové vytvoření dokumentu dostane další volné pořadové označení (`_02`, `_03` atd.), takže existující soubory aplikace nepřepisuje. Ruční úpravy vytvořeného souboru DOCX nelze zpětně načíst do aplikace ani z nich obnovit uloženou zakázku.

## Osnova a předloha

Osnova v `klikatko/catalog.json` vychází z dodaného vzoru. Její právní úplnost nebyla ověřena; před použitím dokumentace zkontrolujte, že odpovídá požadavkům konkrétního záměru.

Aplikace zachovává při tvorbě dokumentu prvky titulní strany předlohy, včetně jejího písma a ukotveného obrázku. Po vytvoření otevřete dokument ve Wordu a vizuálně zkontrolujte sazbu, zalomení a titulní stranu.
