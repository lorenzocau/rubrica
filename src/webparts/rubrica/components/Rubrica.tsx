import * as React from "react";
import { useState, useEffect } from "react";
import { User } from "@microsoft/microsoft-graph-types";

import { getGraphUsers } from "../services/graphService";
import { getVociLista } from "../services/spService";
import { IVociLista, IRubricaUser } from "../types";

import PersonCard from "./PersonCard";
import styles from "./Rubrica.module.scss";
import { User } from "@pnp/graph/users";

function Rubrica() {

  // --- stato ---
  // stato = se valore cambia e la pagina deve aggiornarsi
  // risultati, che cambia in base a quello che riceve e il filtro
  // input, l'input dell'utente per cercare
  // loading, se sta caricando o no
  // errore, se c'è errore o no
  const [risultati, setRisultati] = useState<IRubricaUser[]>([]); // vuota all'inizio
  const [input, setInput] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<boolean>(false);

  // definizione della funzione che carica i dati in risultati
  async function caricaDati() {

    /*  
        se mettiamo await per le due get diventa sequenziale
        invece mettendolo sulla Promise.all sono parallele
        const [,] perchè i valori dopo le call API non cambieranno mai
        quindi sono COSTANTI.
        array [,] (destructuring) per ricevere i risultati dalla Promise.all 
    */
    try {
      const [datiUser, datiLista] = await Promise.all(
        [getGraphUsers(), getVociLista()]
      )
    }
    catch (e) {
      setError(true);
    }
    finally {
      setLoading(false);
    }
    // ora tocca al merge
    // convertiamo datiLista in map



  }
  // --- carica i dati al mount ---
  // mount, quindi devo usare un useEffect con [] alla fine
  useEffect(() => {

  }, []);

  // --- filtra in base alla ricerca ---

  // --- render ---
}

export default Rubrica;

function merge(datiUser: User[], datiLista: IVociLista[]): IRubricaUser[] {
  // dato da ritornare
  var entry: IRubricaUser;
  const rubrica: IRubricaUser[] = [];
  // creiamo la mappa dei dati di sharepoint
  const numeriInterniMap = new Map();
  for (const datoLista of datiLista)
    numeriInterniMap.set(datoLista.mail.toLocaleLowerCase, datoLista.numeroInterno)
  // per ogni User proviamo join con mail, se presente aggiungiamo alla rubrica
  // e rimuoviamo dalla map
  for (const user of datiUser) // tutti gli utenti
  {

    // se presente
    if (numeriInterniMap.get(user.mail?.toLocaleLowerCase) != undefined) {
      entry = { user: user, numeroInterno: numeriInterniMap.get(user.mail?.toLocaleLowerCase) };
      numeriInterniMap.delete(user.mail?.toLowerCase);
    }
    else
      entry = { user: user, numeroInterno: null };
    rubrica.push(entry);
  }
  for (const servizio of datiLista) {
    entry = { user: }
  }


  return rubrica;
}