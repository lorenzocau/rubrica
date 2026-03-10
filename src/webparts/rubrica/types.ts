// Dati da User
import { User } from "@microsoft/microsoft-graph-types";


//---Dati dalla lista SharePoint---
export interface IContattoInterno {
  contatto: string; // email | servizio
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

export type IRubricaEntry = IRubricaUtente | IRubricaServizio;
