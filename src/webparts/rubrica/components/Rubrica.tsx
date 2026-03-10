import * as React from "react";
import { useState, useEffect } from "react";
import { SearchBox } from "@fluentui/react/lib/SearchBox";

// dati mockup
import { getMockGraphUsers as getGraphUsers } from "../mockdata";
import { getMockVociLista as getVociLista } from "../mockdata";
// dati reali
// import { getGraphUsers } from "../services/graphService";
// import { getVociLista } from "../services/spService";
import { IContattoInterno, IRubricaEntry } from "../types";
import PersonCard from "./PersonCard";
import { filtra, merge } from "../rubricaLogic";

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

