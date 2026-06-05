# Ksefik

**Ksefik** to aplikacja do podglądu faktur z KSeF oraz generowania opisu merytorycznego w formacie pliku **Microsoft Word `.docx`**.

Wygenerowany plik `.docx` można wgrać do systemów:

- **EZD RP**
- **EZD PUW**

Projekt zawiera dwie wersje aplikacji:

1. **Wersję Python** — aplikacja `.py`, którą można uruchamiać w Pythonie albo wygenerować z niej plik `.exe`.
2. **Wersję HTML** — plik `ksefik-jst-html.html`, który działa niezależnie bez Pythona, lokalnie w przeglądarce.

---

## Funkcje aplikacji

- podgląd faktury z KSeF,
- wczytywanie pliku faktury KSeF,
- generowanie opisu merytorycznego,
- zapis opisu merytorycznego do pliku `.docx`,
- możliwość użycia wygenerowanego pliku w systemach EZD RP oraz EZD PUW,
- lokalne działanie aplikacji,
- możliwość przygotowania wersji `.exe` z aplikacji Python,
- niezależna wersja HTML działająca bez instalacji Pythona.

---

# Wersja Python

Wersja Python jest przeznaczona dla użytkowników, którzy chcą uruchamiać aplikację z pliku `.py` albo przygotować gotowy plik wykonywalny `.exe` dla systemu Windows.

---

## Wymagania dla wersji Python

Do uruchomienia wersji Python wymagane jest:

- Python 3.10 lub nowszy,
- zainstalowane wymagane biblioteki,
- system Windows, jeżeli ma zostać wygenerowany plik `.exe`.

---

## Uruchomienie aplikacji Python

1. Pobierz repozytorium lub skopiuj pliki projektu na komputer.

2. Otwórz terminal w katalogu projektu.

3. Utwórz środowisko wirtualne:

```bash
python -m venv venv
```

4. Aktywuj środowisko wirtualne.

Dla Windows:

```bash
venv\Scripts\activate
```

Dla Linux/macOS:

```bash
source venv/bin/activate
```

5. Zainstaluj wymagane biblioteki.

Jeżeli w projekcie znajduje się plik `requirements.txt`, użyj:

```bash
pip install -r requirements.txt
```

Jeżeli nie ma pliku `requirements.txt`, zainstaluj biblioteki ręcznie zgodnie z importami użytymi w aplikacji.

6. Uruchom aplikację:

```bash
python ksefik.py
```

> Jeżeli główny plik aplikacji ma inną nazwę niż `ksefik.py`, podmień nazwę pliku w poleceniu.

---

# Generowanie pliku EXE z aplikacji Python

Do przygotowania pliku `.exe` najprościej użyć narzędzia **PyInstaller**.

---

## Instalacja PyInstaller

W aktywnym środowisku wirtualnym wpisz:

```bash
pip install pyinstaller
```

---

## Budowanie pliku EXE

Najprostsza komenda:

```bash
pyinstaller --onefile --windowed ksefik.py
```

Opis parametrów:

- `--onefile` — tworzy jeden plik `.exe`,
- `--windowed` — uruchamia aplikację bez czarnego okna konsoli,
- `ksefik.py` — główny plik aplikacji Python.

Po zakończeniu budowania plik `.exe` znajdziesz w katalogu:

```text
dist/
```

Przykład:

```text
dist/ksefik.exe
```

---

## Budowanie EXE z ikoną

Jeżeli chcesz dodać ikonę aplikacji, przygotuj plik `.ico`, np.:

```text
ikona.ico
```

Następnie użyj komendy:

```bash
pyinstaller --onefile --windowed --icon=ikona.ico ksefik.py
```

---

## Budowanie EXE z dodatkowymi plikami

Jeżeli aplikacja korzysta z dodatkowych plików, np. szablonów `.docx`, grafik, plików konfiguracyjnych albo innych zasobów, należy je dołączyć do kompilacji.

Przykład dla Windows:

```bash
pyinstaller --onefile --windowed --add-data "szablon.docx;." ksefik.py
```

Przykład dla Linux/macOS:

```bash
pyinstaller --onefile --windowed --add-data "szablon.docx:." ksefik.py
```

Uwaga:

- w Windows separatorem jest średnik `;`,
- w Linux/macOS separatorem jest dwukropek `:`.

---

## Czyszczenie plików po budowaniu

Po kompilacji PyInstaller tworzy dodatkowe katalogi i pliki:

```text
build/
dist/
ksefik.spec
```

Najważniejszy jest plik `.exe` w katalogu:

```text
dist/
```

Katalog `build` można usunąć, jeżeli aplikacja została już poprawnie zbudowana.

Plik `ksefik.spec` można zostawić, jeżeli planujesz później budować aplikację z tymi samymi ustawieniami.

---

# Wersja HTML

Projekt zawiera również osobną aplikację:

```text
ksefik-jst-html.html
```

Jest to niezależna wersja aplikacji, która działa lokalnie w przeglądarce internetowej i nie wymaga instalacji Pythona.

Wystarczy skopiować odpowiednie pliki na komputer użytkownika, uruchomić plik HTML, wczytać fakturę z KSeF i wygenerować opis merytoryczny do pliku `.docx`.

---

## Pliki potrzebne do działania wersji HTML

Do uruchomienia wersji HTML potrzebne są:

```text
ksefik-jst-html.html
docx.js
```

Pliki należy trzymać w tym samym katalogu.

Przykład:

```text
Ksefik-HTML/
├── ksefik-jst-html.html
└── docx.js
```

---

## Uruchomienie wersji HTML

1. Skopiuj pliki:

```text
ksefik-jst-html.html
docx.js
```

na komputer użytkownika.

2. Upewnij się, że oba pliki znajdują się w jednym folderze.

3. Kliknij dwukrotnie plik:

```text
ksefik-jst-html.html
```

4. Aplikacja uruchomi się w przeglądarce.

5. Wczytaj fakturę z KSeF.

6. Wygeneruj opis merytoryczny.

7. Zapisz wygenerowany plik `.docx`.

8. Wgraj plik `.docx` do systemu EZD RP lub EZD PUW.

---

# Różnice między wersją Python i HTML

| Funkcja | Wersja Python | Wersja HTML |
|---|---:|---:|
| Podgląd faktury z KSeF | Tak | Tak |
| Generowanie opisu merytorycznego | Tak | Tak |
| Eksport do `.docx` | Tak | Tak |
| Możliwość zrobienia pliku `.exe` | Tak | Nie dotyczy |
| Wymaga Pythona | Tak, przy uruchamianiu `.py` | Nie |
| Działa lokalnie | Tak | Tak |
| Można skopiować na komputer i uruchomić | Tak, po zbudowaniu `.exe` | Tak |

---

# Przykładowy sposób pracy

## Wariant Python / EXE

1. Uruchom aplikację:

```text
ksefik.exe
```

2. Wczytaj fakturę z KSeF.

3. Sprawdź dane faktury w podglądzie.

4. Wygeneruj opis merytoryczny.

5. Zapisz plik `.docx`.

6. Wgraj dokument do EZD RP lub EZD PUW.

---

## Wariant HTML

1. Otwórz plik:

```text
ksefik-jst-html.html
```

2. Wczytaj fakturę z KSeF.

3. Wygeneruj opis merytoryczny.

4. Pobierz plik `.docx`.

5. Wgraj dokument do EZD RP lub EZD PUW.

---

# Struktura projektu

Przykładowa struktura katalogu:

```text
Ksefik/
├── ksefik.py
├── ksefik-jst-html.html
├── docx.js
├── requirements.txt
├── README.md
└── dist/
    └── ksefik.exe
```

---

# Bezpieczeństwo danych

Aplikacja działa lokalnie na komputerze użytkownika.

Wersja HTML nie wymaga serwera i może działać bez wysyłania danych poza komputer użytkownika, o ile pliki są uruchamiane lokalnie.

Faktury z KSeF mogą zawierać dane wrażliwe, dlatego zaleca się:

- przechowywanie plików w bezpiecznej lokalizacji,
- nieudostępnianie faktur osobom nieuprawnionym,
- usuwanie plików tymczasowych po zakończeniu pracy,
- korzystanie z aplikacji na zaufanym komputerze.

---

# Najczęstsze problemy

## Aplikacja Python nie uruchamia się

Sprawdź, czy masz zainstalowanego Pythona:

```bash
python --version
```

Sprawdź, czy zainstalowano wymagane biblioteki:

```bash
pip install -r requirements.txt
```

---

## Plik EXE nie działa na innym komputerze

Spróbuj zbudować aplikację ponownie z opcją:

```bash
pyinstaller --onefile --windowed ksefik.py
```

Jeżeli aplikacja używa dodatkowych plików, trzeba je dołączyć przez `--add-data`.

Przykład:

```bash
pyinstaller --onefile --windowed --add-data "szablon.docx;." ksefik.py
```

---

## Wersja HTML nie generuje pliku DOCX

Sprawdź, czy plik:

```text
docx.js
```

znajduje się w tym samym folderze co:

```text
ksefik-jst-html.html
```

Bez tego biblioteka do generowania plików Word może się nie załadować.

---

# Przeznaczenie

Aplikacja została przygotowana z myślą o jednostkach i użytkownikach pracujących z fakturami z KSeF oraz systemami elektronicznego zarządzania dokumentacją, w szczególności:

- EZD RP,
- EZD PUW,
- jednostki samorządu terytorialnego,
- urzędy,
- działy finansowe,
- kancelarie,
- pracownicy merytoryczni opisujący faktury.

---

# Krótki opis projektu

**Ksefik** to lokalna aplikacja do podglądu faktur z KSeF oraz generowania opisu merytorycznego w pliku `.docx`, gotowego do wykorzystania w systemach EZD RP i EZD PUW. Projekt zawiera wersję Python z możliwością budowy pliku `.exe` oraz niezależną wersję HTML działającą bez instalacji Pythona.

---

# Licencja

Uzupełnij zgodnie z przeznaczeniem projektu, np.:

```text
MIT
```

albo:

```text
Projekt prywatny / wewnętrzny.
```

---

# Autor

Projekt: **Ksefik**

Autor: `Ernest Zając`
