import * as React from "react";
import { useState, useEffect } from "react";
import { User } from "@microsoft/microsoft-graph-types";

import { getGraphUsers } from "../services/graphService";
import { getVociLista } from "../services/spService";
import { IVociLista, IRubricaEntry } from "../types";

import styles from "./Rubrica.module.scss";

function Rubrica() {

  // --- stato ---
  // stato = se valore cambia e la pagina deve aggiornarsi
  // risultati, che cambia in base a quello che riceve e il filtro
  // input, l'input dell'utente per cercare
  // loading, se sta caricando o no
  // errore, se c'è errore o no
  const [rubrica, setRubrica] = useState<IRubricaEntry[]>([]); // vuota all'inizio
  const [input, setInput] = useState<string>("");
  const [filtrati, setFiltrati] = useState<IRubricaEntry[]>([]);
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
      );
      //carichiamo risultati con merge se abbiamo ricevuto
      const merged = merge(datiUser, datiLista);
      setRubrica(merged);
    }
    catch (e) {
      setError(true);
    }
    finally {
      setLoading(false);
    }
  }



  // --- carica i dati al mount ---
  // mount, quindi devo usare un useEffect con [] alla fine
  useEffect(() => {
    caricaDati(); //carica tutti i dati
  }, []);

  // --- filtra in base alla ricerca ---
  useEffect(() => {
    setFiltrati(filtra(input, rubrica));
  }, [input]); // quando cambia l'input


  // --- render ---
  if (loading) return <div>loading...</div>
  if (error) return <div>errore...</div>

  return (
    <>
      <input
        type="text"
        value={input}
        onChange={(e) => setInput(e.target.value)}
        placeholder="Cerca..."
      />
      {filtrati.map((risultato, i) => (
        <div key={i}>
          {risultato.tipo === "utente" ? (
            <>
              <div>{risultato.user.displayName}</div>
              <div>{risultato.user.mail}</div>
              <div>{risultato.user.department}</div>
              <div>{risultato.user.jobTitle}</div>
              <div>{risultato.interno}</div>
            </>
          ) : (
            <>
              <div>{risultato.nome}</div>
              <div>{risultato.interno}</div>
            </>
          )}
        </div>
      ))}
    </>
  );
}

export default Rubrica;

function merge(datiUser: User[], datiLista: IVociLista[]): IRubricaEntry[] {
  // dato da ritornare
  const rubrica: IRubricaEntry[] = [];
  // creiamo la mappa dei dati di sharepoint
  const internimap = new Map();

  let entry: IRubricaEntry;
  for (const datoLista of datiLista)
    internimap.set(datoLista.contatto.toLowerCase(), datoLista.interno)
  // per ogni User proviamo join con mail, se presente aggiungiamo alla rubrica
  // e rimuoviamo dalla map
  for (const user of datiUser) // tutti gli utenti
  {
    if (internimap.get(user.mail?.toLowerCase()) != undefined) {
      entry = { tipo: "utente", user: user, interno: internimap.get(user.mail?.toLowerCase()) };
      internimap.delete(user.mail?.toLowerCase());
    }
    else 
    {
      entry = { tipo: "utente", user: user, interno: null };
    }
    rubrica.push(entry);
  }
  // ricostruisco array dalla mappa...
  for (const [contatto, interno] of Array.from(internimap)) {
    entry = { tipo: "servizio", nome: contatto, interno: interno };
    rubrica.push(entry);
  }



  return rubrica;
}

function filtra(input: string, risultati: IRubricaEntry[]): IRubricaEntry[] {
  if (input.trim() == "") return []; // no input -> no risultati

  const q = input.toLowerCase();
  const filtrati: IRubricaEntry[] = [];

  for (const entry of risultati) {
    if (entry.tipo === "utente") {
      if (entry.user.displayName?.toLowerCase().includes(q) ||
        entry.user.mail?.toLowerCase().includes(q) ||
        entry.user.department?.toLowerCase().includes(q) ||
        entry.interno?.toLowerCase().includes(q))

        filtrati.push(entry);
    }
    else {
      if (entry.nome.toLowerCase().includes(q) ||
        entry.interno.toLowerCase().includes(q))

        filtrati.push(entry);
    }
  }
  return filtrati;
}
