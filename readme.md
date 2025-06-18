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
## Tworzenie formularza
Formularz musi zawierać takie pytania:

![image](https://github.com/user-attachments/assets/7ccb19cf-55c9-446c-95e2-c06a0a424852)

Oraz koniecznie musi być w micrpsoft forms i mieć takie ustawienia

![image](https://github.com/user-attachments/assets/b30d90aa-50fc-4ca9-bf25-44e478b7cab4)

Po zebraniu danych jakie nam potrzeba od członków koła trzeba te dane wyeksportować, w tym celu przechodzimy do zakładki wyświetl odpowiedzi

![image](https://github.com/user-attachments/assets/1c4987ea-a74a-4b2d-8fa4-b676c3e0c048)

Następnie po prawej stronie mamy excela w ktorego należy wejść

![image](https://github.com/user-attachments/assets/34f64e1a-4b6c-4db6-bfc1-626ec87beb19)

Excel powinien zawierać takie kolumny

![image](https://github.com/user-attachments/assets/6da6c194-6d6f-44f9-9d9b-979282f1ec2b)


Aby teraz wyeksportować dane jako scv należy kliknąć file > export > DOwnload as CSV UTF-8

![image](https://github.com/user-attachments/assets/2b77c20f-3864-4de6-8702-7bc0bef47454)

Zapisujemy gdzieś gdzie będzie wygodnie to znaleźć i możemu uruchamiać program

---

## Uruchomienie
Przed uruchomieniem upewnij sie że w pliku szablon.docx wpisałeś nazwe swojego koła

```bash
python main.py sciezka_do_pliku.csv
```

---

## Uwagi
Nazwy kolumn są bardzo ważne ponieważ na ich pdstawie skrypt wie co gdzie wstawić więc jeżeli nie są zgodne z tym co pokazałem wcześniej w excelu należy je zmienić
