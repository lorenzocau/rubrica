// Dati da User
import { User } from "@microsoft/microsoft-graph-types";

// User di Graph + faxNumber: in AD contiene il "numero breve" del cellulare aziendale.
// Graph lo restituisce solo se richiesto con $select (vedi graphService).
// L'intersezione funziona sia se la versione dei tipi di Graph lo dichiara già, sia se no.
export type UtenteGraph = User & { faxNumber?: string | null };


//---Dati dalla lista SharePoint---
export interface IContattoInterno {
  contatto: string; // email | servizio
  interno: string;
  reparto?: string; // colonna "Reparto" (opzionale): reparto del servizio
}

// Riga della lista "Esterni" (fornitori, clienti, consulenti...)
export interface IContattoEsterno {
  nome: string;        // colonna "Nome"
  numero: string;      // colonna "Numero": numero per intero ("" se assente)
  numeroBreve: string; // colonna "Numero breve" ("" se assente)
}


export interface IRubricaUtente {
  tipo: "utente";   // discriminante
  user: UtenteGraph;
  interno: string | null;
}

export interface IRubricaServizio {
  tipo: "servizio";
  nome: string;     // discriminante
  interno: string;
  reparto: string;  // "" se la lista non ha il reparto per questa voce
}

// Contatto esterno (fornitori, consulenti...): viene dalla lista "Esterni",
// non viene mai abbinato a un utente Graph.
// Se un contatto ha più numeri, sono uniti con " / " (vedi SEPARATORE_INTERNI).
export interface IRubricaEsterno {
  tipo: "esterno";
  nome: string;
  numero: string;       // numero per intero ("" se assente)
  numeroBreve: string;  // numero breve ("" se assente)
}

export type IRubricaEntry = IRubricaUtente | IRubricaServizio | IRubricaEsterno;


//---Impostazioni del web part (pannello proprietà)---
export interface IRubricaWebPartProps {
  // Generale
  titolo: string;
  placeholderRicerca: string;
  risultatiPerPagina: number;
  mostraElencoCompleto: boolean;   // true = senza ricerca mostra tutta la rubrica

  // Liste SharePoint: titolo della lista oppure percorso (es. /sites/intranet/Lists/Esterni)
  listaServizi: string;
  listaEsterni: string;

  // Esterni
  mostraEsterni: boolean;               // abilita la funzione (carica la lista + mostra la checkbox)
  esterniInizialmenteVisibili: boolean; // stato iniziale della checkbox

  // Filtri sugli account Graph
  soloMembri: boolean;            // true = nasconde gli ospiti (userType Guest)
  nascondiSenzaMail: boolean;     // true = nasconde gli account senza mail
  nascondiSenzaTelefono: boolean; // true = nasconde gli utenti senza alcun numero (interno, cellulare, cellulare breve)

  // Visualizzazione
  mostraContatore: boolean;
  mostraPresenza: boolean;
  mostraMail: boolean;
  mostraReparto: boolean;     // job title degli utenti / reparto dei servizi
  mostraCellulare: boolean;   // cellulare aziendale + numero breve
  linkCliccabili: boolean;    // mailto, tel e chat Teams

  // Sviluppo/test
  usaDatiMock: boolean;
}

// Sottoinsieme usato da PersonCard
export interface IOpzioniCard {
  mostraPresenza: boolean;
  mostraMail: boolean;
  mostraReparto: boolean;
  mostraCellulare: boolean;
  linkCliccabili: boolean;
}
