import * as React from "react";
import { useState, useEffect } from "react";
import { User } from "@microsoft/microsoft-graph-types";
import { SearchBox } from "@fluentui/react/lib/SearchBox";

// dati mockup
import { getMockGraphUsers as getGraphUsers } from "../mockdata";
import { getMockVociLista as getVociLista } from "../mockdata";
// dati reali
// import { getGraphUsers } from "../services/graphService";
// import { getVociLista } from "../services/spService";
import { IVociLista, IRubricaEntry } from "../types";
import PersonCard from "./PersonCard";

function Rubrica() {

  // --- stato ---
  const [rubrica, setRubrica] = useState<IRubricaEntry[]>([]);
  const [input, setInput] = useState<string>("");
  const [filtrati, setFiltrati] = useState<IRubricaEntry[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [warning, setWarning] = useState<string>("");

  async function caricaDati() {
    try {
      const [datiGraph, datiLista] = await Promise.all([getGraphUsers(), getVociLista()]);
      if (datiGraph.length === 0) setWarning("Problema con Microsoft Graph");
      if (datiLista.length === 0) setWarning("Problema con la lista Contatti");
      const merged = merge(datiGraph, datiLista);
      setRubrica(merged);
    }
    // errori gestiti nei servizi
    finally {
      setLoading(false);
    }
  }

  // --- carica i dati al mount ---
  useEffect(() => {
    caricaDati();
  }, []);

  // --- filtra in base alla ricerca ---
  useEffect(() => {
    setFiltrati(filtra(input, rubrica));
  }, [input]);

  // --- render ---
  if (loading) return <div>Caricamento...</div>;

  return (
    <>
      {warning && <div className={warning}>{warning}</div>}
      <SearchBox
        placeholder="Cerca per nome, reparto, interno..."
        value={input}
        onChange={(event, text) => setInput(text ?? "")}
        onClear={() => setInput("")}
      />
      {filtrati.slice(0, 30).map((entry, i) => (
        <PersonCard key={i} entry={entry} />
      ))}
    </>
  );
}

export default Rubrica;

function merge(datiGraph: User[], datiLista: IVociLista[]): IRubricaEntry[] {
  const rubrica: IRubricaEntry[] = [];
  const internimap = new Map();

  let entry: IRubricaEntry;
  for (const datoLista of datiLista)
    internimap.set(datoLista.contatto.toLowerCase(), datoLista.interno);

  for (const user of datiGraph) {
    if (internimap.get(user.mail?.toLowerCase()) != undefined) {
      entry = { tipo: "utente", user: user, interno: internimap.get(user.mail?.toLowerCase()) };
      internimap.delete(user.mail?.toLowerCase());
    }
    else {
      entry = { tipo: "utente", user: user, interno: null };
    }
    rubrica.push(entry);
  }

  for (const [contatto, interno] of Array.from(internimap)) {
    entry = { tipo: "servizio", nome: contatto, interno: interno };
    rubrica.push(entry);
  }

  return rubrica;
}

function filtra(input: string, risultati: IRubricaEntry[]): IRubricaEntry[] {
  if (input.trim() === "") return []; // nulla mostrato finché non si cerca

  const q = input.toLowerCase();
  const filtrati: IRubricaEntry[] = [];

  for (const entry of risultati) {
    if (entry.tipo === "utente") {
      if (entry.user.displayName?.toLowerCase().includes(q) ||
        entry.user.mail?.toLowerCase().includes(q) ||
        entry.user.department?.toLowerCase().includes(q) ||
        entry.interno?.toLowerCase().includes(q) ||
        entry.user.jobTitle?.toLowerCase().includes(q))

        filtrati.push(entry);
    }
    else {
      if (entry.nome.toLowerCase().includes(q) ||
        entry.interno.toLowerCase().includes(q) ||
        "servizio".includes(q))

        filtrati.push(entry);
    }
  }
  return filtrati;
}