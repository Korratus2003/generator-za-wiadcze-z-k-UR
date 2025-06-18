# Generator Zaświadczeń w formacie DOCX

Automatyczny generator spersonalizowanych zaświadczeń na podstawie danych zawartych w pliku CSV, z wykorzystaniem szablonu DOCX.

---

## Opis działania

Skrypt wczytuje dane uczestników z pliku CSV, przetwarza je i generuje zaświadczenia w formacie `.docx`, podstawiając odpowiednie dane w przygotowanym szablonie dokumentu.
Następnie zapisuje plik w folderze zaswiadzcenia w formacie "Imie Nazwisko zaświadczenie.docx"

---

## Funkcjonalności

- Parsowanie danych osobowych i identyfikatorów z adresu e-mail.
- Obsługa różnych formatów dat.
- Wstawianie imienia i nazwiska, numeru albumu, kierunku studiów, dat uczestnictwa oraz osiągnięć do dokumentu.
- Automatyczne tworzenie katalogu `zaswiadczenia` i zapisywanie wygenerowanych plików.
- Walidacja poprawności dat oraz istnienia pliku CSV.

---

## Wymagania

Projekt wymaga zainstalowanych bibliotek Python, które można łatwo zainstalować za pomocą pliku `requirements.txt`.

Aby zainstalować wszystkie zależności, wykonaj polecenie:

```bash
pip install -r requirements.txt
```

---

## Uruchomienie

```bash
python main.py sciezka_do_pliku.csv
```