# Ksefik

**Ksefik** to aplikacja do podglądu faktur z KSeF oraz generowania opisu merytorycznego do pliku **Microsoft Word `.docx`**, który można następnie wgrać do systemów:

- **EZD RP**
- **EZD PUW**

Projekt zawiera dwie wersje aplikacji:

1. **Aplikację Python `.py`**  
   Wersja desktopowa uruchamiana z pliku Python, którą można skompilować do pliku `.exe`.

2. **Aplikację HTML `ksefik-jst-html.html`**  
   Osobna, niezależna wersja działająca bez Pythona. Wystarczy skopiować pliki na komputer i uruchomić aplikację w przeglądarce.

---

## Do czego służy Ksefik?

Aplikacja umożliwia wczytanie faktury pobranej z KSeF, podgląd jej danych oraz wygenerowanie gotowego opisu merytorycznego w formacie `.docx`.

Wygenerowany plik Word można wykorzystać jako załącznik lub dokument pomocniczy w obiegu dokumentów w systemach **EZD RP** oraz **EZD PUW**.

---

## Główne funkcje

- wczytywanie faktury z KSeF,
- podgląd danych faktury,
- przygotowanie opisu merytorycznego,
- generowanie pliku Microsoft Word `.docx`,
- możliwość wykorzystania wygenerowanego pliku w EZD RP i EZD PUW,
- wersja Python możliwa do spakowania jako aplikacja `.exe`,
- wersja HTML działająca lokalnie, bez instalacji Pythona.

---

# Wersja Python

## Wymagania

Do uruchomienia wersji Python wymagane jest:

- Python 3.10 lub nowszy,
- zainstalowane wymagane biblioteki,
- system Windows, jeżeli planujesz wygenerować plik `.exe`.

---

## Uruchomienie aplikacji Python

1. Pobierz repozytorium lub skopiuj pliki projektu na komputer.

2. Otwórz terminal w katalogu projektu.

3. Utwórz środowisko wirtualne:

```bash
python -m venv venv
