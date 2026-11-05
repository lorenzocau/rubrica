// Dati da User
import { User } from "@microsoft/microsoft-graph-types";


//---Dati dalla lista SharePoint---
export interface IContattoInterno {
  contatto: string; // email | servizio | esterno
  interno: string;
}


export interface IRubricaUtente {
  tipo: "utente";   // discriminante
  user: User;
  interno: string | null;
}

export interface IRubricaServizio {
  tipo: "servizio";
  nome: string;     // discriminante
  interno: string;
}

// Contatto esterno (fornitori, consulenti...): viene dalla lista "Esterni",
// ha la stessa struttura di un servizio ma non viene mai abbinato a un utente Graph.
export interface IRubricaEsterno {
  tipo: "esterno";
  nome: string;
  interno: string;
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
  soloMembri: boolean;        // true = nasconde gli ospiti (userType Guest)
  nascondiSenzaMail: boolean; // true = nasconde gli account senza mail

  // Visualizzazione
  mostraContatore: boolean;
  mostraPresenza: boolean;
  mostraMail: boolean;
  mostraReparto: boolean;
  mostraCellulare: boolean;
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
