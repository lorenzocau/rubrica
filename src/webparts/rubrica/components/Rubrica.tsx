import * as React from "react";
import { useState, useEffect } from "react";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { Text } from "@fluentui/react/lib/Text";

// dati mockup
import { getMockGraphUsers as getGraphUsers } from "../mockdata";
import { getMockVociLista as getVociLista, LIST_NAME} from "../mockdata";
// dati reali
//import { getGraphUsers } from "../services/graphService";
//import { getVociLista, LIST_NAME } from "../services/spService";
import { IRubricaEntry } from "../types";
import PersonCard from "./PersonCard";
import { filtra, merge } from "../rubricaLogic";
import { MessageBar, MessageBarType } from "@fluentui/react";

const PAGE_SIZE = 30;

// Esegue una sorgente dati: se fallisce registra l'avviso e restituisce [],
// così l'altra sorgente viene comunque mostrata.
async function conAvviso<T>(promessa: Promise<T[]>, avviso: string, avvisi: string[]): Promise<T[]> {
  try {
    return await promessa;
  } catch (e) {
    console.warn(avviso, e);
    avvisi.push(avviso);
    return [];
  }
}

function Rubrica() {
  const [rubrica, setRubrica] = useState<IRubricaEntry[]>([]);
  const [ricerca, setRicerca] = useState<string>("");
  const [filtrati, setFiltrati] = useState<IRubricaEntry[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [avvisi, setAvvisi] = useState<string[]>([]);
  const [pagina, setPagina] = useState<number>(0);

  async function caricaDati() {
    const avvisiCaricamento: string[] = [];

    try {
      const [datiGraph, datiLista] = await Promise.all([
        conAvviso(
          getGraphUsers(),
          "Recupero utenti da Microsoft Graph non riuscito: la rubrica potrebbe essere incompleta.",
          avvisiCaricamento
        ),
        conAvviso(
          getVociLista(),
          `Lista "${LIST_NAME}" non trovata o non accessibile in questo sito: gli interni non sono disponibili.`,
          avvisiCaricamento
        ),
      ]);
      setRubrica(merge(datiGraph, datiLista));
    } catch (e) {
      console.error("Errore imprevisto nel caricamento della rubrica.", e);
      avvisiCaricamento.push("Errore imprevisto nel caricamento della rubrica.");
    } finally {
      setAvvisi(avvisiCaricamento);
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
      {avvisi.map((avviso, i) => (
        <MessageBar key={i} messageBarType={MessageBarType.warning}>{avviso}</MessageBar>
      ))}

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