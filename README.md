# FlatScanner

Skaner ofert wynajmu mieszkan w Gdansku. Przeszukuje OLX i Otodom, wyciaga szczegolowe dane (cena, koszty ukryte, powierzchnia, udogodnienia) i zapisuje wyniki do CSV/XLSX ze zdjeciami.

## Funkcje

- Przeszukiwanie OLX.pl i Otodom.pl
- Nowoczesny panel GUI (CustomTkinter) z wyborem kryteriow
- Ekstrakcja ukrytych kosztow z opisow (czynsz admin, media, prad, gaz, parking, itp.)
- Zdjecia osadzone w pliku XLSX
- Klikalne linki do ofert
- TOP 10 ofert wg kryteriow uzytkownika
- Statystyki na zywo podczas skanowania

## Instalacja

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

## Uruchomienie

### GUI (zalecane)
```bash
python olx_gui.py
```

### Tylko scraper (CLI)
```bash
python olx_scraper.py
```

## Wymagania

- Python 3.10+
- Google Chrome zainstalowany w systemie
- Windows 10/11
