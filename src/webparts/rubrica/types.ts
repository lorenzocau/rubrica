import { User } from "@microsoft/microsoft-graph-types";

//---Dati dalla lista SharePoint---
export interface IVociLista {
  mail: string; // email | servizio
  numeroInterno: string;
}

// ---modello finale (SP + Graph se presente)---
export interface IRubricaUser {
  user: User | null;  // null se servizio
  numeroInterno: string | null; // null se num interno non presente
}