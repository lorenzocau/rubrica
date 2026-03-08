// Dati da User
import { User } from "@microsoft/microsoft-graph-types";


//---Dati dalla lista SharePoint---
export interface IVociLista {
  contatto: string; // email | servizio
  interno: string;
}

// ---modello finale con discriminated union tramite tipo
// per evitare tutti i possibili null in più
export interface IRubricaUtente {
  tipo: "utente";
  user: User;
  interno: string | null;
}

export interface IRubricaServizio {
  tipo: "servizio";
  nome: string;  // non più nullable
  interno: string;
}

export type IRubricaEntry = IRubricaUtente | IRubricaServizio;
