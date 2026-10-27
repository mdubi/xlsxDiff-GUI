# xlsxDiff GUI 1.2.1 — instrukcja (PL)

Graficzna aplikacja do porównywania dwóch skoroszytów XLSX lub XLSM na Windows, z
kolejką porównań. **Wersja zmodyfikowana** narzędzia [xlsxDiff](https://github.com/rafal-dot/xlsxDiff)
autorstwa **Rafała Czeczótki** (© 2020–2023). Licencja całej pochodnej:
**GNU AGPL-3.0-or-later**. Za rozwój wersji GUI (2026) odpowiada
**Michał Dubrawski**, przy wsparciu programistycznym i dokumentacyjnym
**ChatGPT (OpenAI GPT-6)**. Zobacz `LICENSE`, `NOTICE.txt`, `SOURCE.md`.

## Jak działa

1. Wybierz **starszą** i **nowszą** wersję pliku XLSX lub XLSM (można łączyć oba formaty).
2. Wskaż plik raportu, np. `DIFF_ankieta_2025_vs_2026.xlsx`.
3. Kliknij **Dodaj do kolejki**; można dodać wiele par plików.
4. W opcjach możesz włączyć porównywanie formuł, indeksy wierszy/kolumn itd.
5. Uruchom porównanie. Raport jest oddzielnym XLSX z oznaczeniem zmian. Makra VBA nie są porównywane, kopiowane ani uruchamiane.

W raporcie: czerwony/przekreślony fragment został usunięty, niebieski/
podkreślony został dodany, niezmienione komórki mają szare tło; zakładki
zmienione, dodane i usunięte są odpowiednio białe, niebieskie i czerwone.
Kolor zakładki nie stanowi nieomylnego dowodu braku zmiany, dlatego warto
zajrzeć także do szarych zakładek. Raport nie jest wierną kopią formatowania
skoroszytu. **Program nie oblicza formuł Excel**: w trybie wartości porównuje
oznaczone w pliku wyniki ostatniego przeliczenia.

Nowa opcja „Zmiany w liczbach pokazuj w całości” (w GUI domyślnie aktywna):
zamiast przekreślenia tylko zmienionej cyfry program traktuje całą liczbę
lub tekstowy procent jako usunięty i całą nową wartość jako dodaną.

## Zmiana języka

W prawym górnym rogu jest **Język / Language**: `Polski` albo `English`.
Zmiana działa bez restartu i nie kasuje kolejki ani opcji. Języka nie
przełącza się w czasie wykonywania zadań.

Dostępne są dwa rodzaje EXE: jeden uruchamia się domyślnie po polsku, drugi
po angielsku. **Każdy EXE oferuje oba języki**. Język wybrany w GUI nie jest
zapisywany jako ustawienie globalne, dlatego przy ponownym uruchomieniu
program wraca do domyślnego języka danego builda.

## Przygotowanie środowiska Windows (CMD, Python 3.12–3.14)

Wypakuj ZIP do katalogu bez konieczności uprawnień administratora, np. do
`C:\Users\...\Documents\XlsxDiff_GUI`. Otwórz CMD w katalogu projektu.
Nie musisz mieć komendy `py`: użyj działającego polecenia `python`.

```bat
python --version
python -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

### Start bez EXE

```bat
.venv\Scripts\python.exe xlsxDiff_GUI.py --lang pl
.venv\Scripts\python.exe xlsxDiff_GUI.py --lang en
```

### Budowanie EXE z ikoną i domyślnym językiem

```bat
build_onefile.cmd pl
build_onefile.cmd en
```

Wyniki w `dist`:
- `XlsxDiff_GUI_PL.exe` — domyślny polski, przełącznik na angielski;
- `XlsxDiff_GUI_EN.exe` — domyślny angielski, przełącznik na polski.

Kompilacja katalogowa (często szybciej uruchamia się i jest wygodniejsza
w środowisku firmowym):

```bat
build_onedir.cmd pl
build_onedir.cmd en
```

Wyniki: katalogi `dist\XlsxDiff_GUI_PL\` i `dist\XlsxDiff_GUI_EN\`.
Przy dystrybucji `onedir` przekaż **cały folder**, nie tylko znajdujący się
w nim EXE.

Skrypty używają `--icon assets\xlsxDiff.ico` oraz pakują PNG/ICO do GUI.
Nie wymagają `py`, ale wymagają pakietów z requirements i dostępu do plików
w katalogu projektu. **Nie tworzyłem tutaj gotowego Windows EXE** — trzeba
go skompilować i uruchomić na Windows; wydania PL i EN należy przetestować.

Jeżeli konsola zgłasza brak `python`, ustal lokalizację przez `where python`.
Jeśli `python -m venv .venv` zadziała, nie instaluj drugiej wersji Pythona.

## Stary interfejs CLI

```bat
.venv\Scripts\python.exe xlsxDiff.py "stary.xlsx" "nowy.xlsx" "DIFF.xlsx" -N
```

Flagi: `-f` porównuj formuły; `-x` nagłówki; `-c` indeksowanie wierszy wg
kolumn; `-r` indeksowanie kolumn wg wierszy; `-X` wyłącz kolorowanie całych
wstawianych/usuwanych wierszy; `-a` filtr; `-e` pomiń puste komórki;
`-N` porównuj całe liczby; `-v` szczegółowy postęp. W GUI tryb `-N` jest
włączony domyślnie, w CLI — dla zgodności — nadal trzeba go jawnie dodać.

## Licencja i publikacja na GitHub

1. Utwórz publiczne repozytorium `mdubi/xlsxDiff-GUI` (puste, bez
   automatycznego README/licencji), z opisem:
   `Bilingual Excel XLSX comparison GUI based on xlsxDiff by Rafał Czeczótka`.
2. W `project_info.py` ustawiony jest planowany adres
   `https://github.com/mdubi/xlsxDiff-GUI`. **Sprawdź, czy repozytorium jest
   już publiczne i adres działa PRZED zbudowaniem EXE do dystrybucji.**
   W oknie **O programie → Kod źródłowy** otworzy się ten adres.
3. Zachowaj oryginalne autorstwo, licencję AGPL, `NOTICE.txt` oraz oryginalny
   kod `upstream_original/xlsxDiff.py`. Nie usuwaj informacji o modyfikacjach.
4. Przy wydaniu `v1.2.1` publikuj odpowiadający kod źródłowy i pakiety EXE
   zbudowane z **tego samego commitu**. Dodaj w wydaniu archiwum źródeł,
   pliki `LICENSE`, `NOTICE.txt`, `SOURCE.md` i odpowiednie notices bibliotek.
5. Nie publikuj firmowych skoroszytów, raportów ani `.venv/`, prywatnych
   ścieżek, logów, kluczy lub danych klientów. Zobacz `.gitignore` i
   `RELEASE_CHECKLIST.md`, `THIRD_PARTY.md`.

Przykładowe komendy Git (po stworzeniu pustego repozytorium na GitHub):

```bat
git init
git branch -M main
git add .
git commit -m "Release bilingual xlsxDiff GUI 1.2.1"
git remote add origin https://github.com/mdubi/xlsxDiff-GUI.git
git push -u origin main
git tag v1.2.1
git push origin v1.2.1
```

Utwórz następnie ZIP tylko z publicznym kodem (bez przypadkowych plików klientów):

```bat
package_source.cmd
```

Powstanie `XlsxDiff_GUI_v1.2.1_SOURCE.zip`. **Sprawdź jego zawartość** przed
przekazaniem poza firmę. Następnie na GitHub wybierz **Releases → Draft a new release**, ustaw
`v1.2.1`, dodaj EXE i archiwum kompletnego kodu tej samej wersji.
Uwaga: GitHub sam dostarcza pobieranie źródeł dla tagu, ale dla prostoty
i audytowalności zalecam dołączyć dodatkowy ZIP z kompletnej paczki.

AGPL nie nakłada automatycznie obowiązku publikacji kodu na GitHubie, lecz
przy dystrybucji binariów należy zapewnić zgodny dostęp do Corresponding
Source. Jeśli chcesz rozprowadzać zamknięte wersje programu, wymagana byłaby
inna zgoda właścicieli praw (AGPL nie pozwala po prostu utajnić modyfikacji).

**To praktyczna instrukcja, a nie opinia prawna.**
