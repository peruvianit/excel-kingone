# 🎾 KingOne Padel Tournament Generator

**KingOne Padel Tournament Generator** è uno script Python pensato per organizzatori di eventi sportivi, progettato per semplificare e ottimizzare la gestione di mini tornei di padel direttamente in **Microsoft Excel**.

## 🧠 Descrizione del Progetto

Questo strumento genera automaticamente un file Excel professionale con:
- ✔️ Elenco dei giocatori divisi per posizione (destro/sinistro)
- ✔️ Programmazione delle partite su più campi in turni ottimizzati
- ✔️ Calcolo dei games vinti da ciascun giocatore
- ✔️ Collegamenti dinamici tra i fogli per aggiornamenti automatici
- ✔️ Protezione dei dati con password

Il file Excel generato include 3 fogli:
1. **Giocatori** – elenco completo con posizioni (D/S)
2. **Partite** – planning dettagliato dei turni, squadre e punteggi
3. **Riepilogo** – classifica dei games vinti per giocatore

## ⚙️ Parametri Configurabili

Prima dell’esecuzione, l’utente può impostare:
- ⏳ Durata complessiva del torneo (in minuti)
- 🕓 Durata di ogni partita (in minuti)
- ⌚ Orario d’inizio del torneo
- 🏟️ Numero di campi disponibili
- 👥 Numero di giocatori partecipanti
- 📂 Percorso di esportazione del file Excel

## 🔎 Requisiti e Validazioni

- Il numero di giocatori **deve essere multiplo di 4**
- Non può eccedere la capacità massima dei campi per ciascun turno
- I giocatori sono divisi in **destri (D)** e **sinistri (S)**, accoppiati in squadre miste

## ♻️ Algoritmo di Ottimizzazione

L’algoritmo cerca di:
- Garantire che **tutti i giocatori giochino ad ogni turno**
- Evitare che **gli stessi giocatori si incontrino troppo spesso**
- Alternare le combinazioni tra destri e sinistri nei vari turni

## 📤 Output Excel

- Formule collegate: i nomi nei fogli Partite e Riepilogo sono **collegati al foglio Giocatori**

## 🧪 Come Usarlo

1. Assicurati di avere Python installato (>= 3.8)
2. Installa le dipendenze:
   ```bash
   pip install openpyxl pandas
3. Esecuazione
 Installa le dipendenze:
   ```bash
      Params : duration_minutes, match_duration, start_time, num_courts, num_players, export_path
      
      python kingone-generator.py 120 20 18:00 4 16 ./torneo_padel_V2.xlsx   
