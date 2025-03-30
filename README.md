# Analiza stanu magazynu

**Aplikacja desktopowa do analizy stanów magazynowych na podstawie plików Excel.**

## Opis

Program umożliwia zbiorczą analizę danych magazynowych zapisanych w arkuszach Excel, w których znajdują się informacje o zakupach dokonanych w danym miesiącu. Na podstawie danych wejściowych aplikacja tworzy dwie tabele:

- **Szczegóły:** wszystkie pozycje ze stanami > 0.
- **Suma:** podsumowanie pozycji według `Index`, `Rozmiar`, `Towar`, wraz z sumą ilości, wartości zakupu i sprzedaży.

Na dole aplikacji widoczne jest zbiorcze podsumowanie: całkowita ilość, łączna wartość zakupu i sprzedaży.

## Kluczowe funkcje

- Obsługa wielu plików Excel jednocześnie.
- Grupowanie i sumowanie duplikujących się pozycji.
- Wyświetlanie danych w dwóch przejrzystych tabelach.
- Eksport danych do pliku `.xlsx` z dwiema zakładkami: `Szczegóły` i `Suma`.
- Automatyczne dostosowanie szerokości kolumn.
- Wyświetlanie cen z dokładnością do dwóch miejsc po przecinku.
- Interfejs w języku polskim.

## Wymagania

- Python 3.7+
- Pakiety:
  - `tkinter`
  - `openpyxl`

## Uruchomienie

```bash
python inventory_app.py
```

## Generowanie pliku EXE

Generowania EXE przy pomocy pakietu cx_Freeze z poziomu konsoli Pycharm: 

```python setup.py build```

