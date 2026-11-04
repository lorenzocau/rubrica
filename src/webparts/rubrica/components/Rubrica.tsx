import * as React from "react";
import { useState, useEffect } from "react";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { Text } from "@fluentui/react/lib/Text";

// dati mockup
// import { getMockGraphUsers as getGraphUsers } from "../mockdata";
// import { getMockVociLista as getVociLista } from "../mockdata";
// dati reali
import { getGraphUsers } from "../services/graphService";
import { getVociLista } from "../services/spService";
import { IRubricaEntry } from "../types";
import PersonCard from "./PersonCard";
import { filtra, merge } from "../rubricaLogic";
import { MessageBar, MessageBarType } from "@fluentui/react";

const PAGE_SIZE = 30;

function Rubrica() {
  const [rubrica, setRubrica] = useState<IRubricaEntry[]>([]);
  const [ricerca, setRicerca] = useState<string>("");
  const [filtrati, setFiltrati] = useState<IRubricaEntry[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [avviso, setAvviso] = useState<string>("");
  const [pagina, setPagina] = useState<number>(0);

  async function caricaDati() {
    try {
      const [datiGraph, datiLista] = await Promise.all([getGraphUsers(), getVociLista()]);
      if (datiGraph.length === 0) setAvviso("Recupero utenti da Microsoft Graph non riuscito.");
      if (datiLista.length === 0) setAvviso("Recupero voci dalla lista Contatti non riuscito.");
      setRubrica(merge(datiGraph, datiLista));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => { void caricaDati(); }, []);

  useEffect(() => {
    setFiltrati(filtra(ricerca, rubrica));
    setPagina(0); // reset alla prima pagina ad ogni ricerca
  }, [ricerca, rubrica]);

  if (loading) return <div>Caricamento...</div>;

  const utenti = rubrica.filter(e => e.tipo === "utente").length;
  const servizi = rubrica.filter(e => e.tipo === "servizio").length;

  const inizio = pagina * PAGE_SIZE;
  const paginaCorr = filtrati.slice(inizio, inizio + PAGE_SIZE);
  const totPagine = Math.ceil(filtrati.length / PAGE_SIZE);

  return (
    <>
      {avviso && <MessageBar messageBarType={MessageBarType.warning}>{avviso}</MessageBar>}

      {/* ── Contatore caricati ── */}
      <Text variant="small" style={{ color: "#605e5c", display: "block", marginBottom: 8 }}>
        Rubrica: <strong>{utenti}</strong> utenti · <strong>{servizi}</strong> servizi
      </Text>

      <SearchBox
        placeholder="Cerca per nome, reparto, interno..."
        value={ricerca}
        onChange={(_e, text) => setRicerca(text ?? "")}
        onClear={() => setRicerca("")}
      />

      {/* ── Contatore risultati + pagina ── */}
      {ricerca.trim() !== "" && (
        <Text variant="small" style={{ color: "#605e5c", display: "block", margin: "6px 0" }}>
          {filtrati.length} risultat{filtrati.length === 1 ? "o" : "i"}
          {totPagine > 1 && ` — pagina ${pagina + 1} di ${totPagine}`}
        </Text>
      )}

      {paginaCorr.map((entry, i) => <PersonCard key={i} entry={entry} />)}

      {/* ── Paginazione ── */}
      {totPagine > 1 && (
        <div style={{ display: "flex", gap: 8, padding: "12px 0", justifyContent: "center" }}>
          <DefaultButton
            disabled={pagina === 0}
            onClick={() => setPagina(p => p - 1)}
          >← Precedente</DefaultButton>
          <DefaultButton
            disabled={pagina >= totPagine - 1}
            onClick={() => setPagina(p => p + 1)}
          >Successiva →</DefaultButton>
        </div>
      )}
    </>
  );
}

export default Rubrica;