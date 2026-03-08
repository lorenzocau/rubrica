import * as React from "react";
import { useState, useEffect } from "react";
import { User } from "@microsoft/microsoft-graph-types";

import { getGraphUsers } from "../services/graphService";
import { getVociLista } from "../services/spService";
import { IVociLista, IRubricaUser } from "../types";

import PersonCard from "./PersonCard";
import styles from "./Rubrica.module.scss";

const Rubrica: React.FC = () => {
  const [rubrica, setRubrica] = useState<IRubricaUser[]>([]);
  const [filtrata, setFiltrata] = useState<IRubricaUser[]>([]);
  const [cerca, setCerca] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(true);
  const [errore, setErrore] = useState<string | null>(null);

  // Al mount carica i dati da Graph e SP in parallelo, poi fa il merge
  useEffect(() => {
    const caricaDati = async (): Promise<void> => {
      try {
        const [utentiGraph, vociSP] = await Promise.all([
          getGraphUsers(),
          getVociLista(),
        ]);
        const merged = merge(utentiGraph, vociSP);
        setRubrica(merged);
        setFiltrata(merged);
      } catch (e) {
        setErrore("Errore durante il caricamento della rubrica.");
      } finally {
        setLoading(false);
      }
    };

    caricaDati();
  }, []);

  // Aggiorna la lista filtrata ogni volta che cambia il testo di ricerca
  useEffect(() => {
    if (!cerca.trim()) {
      setFiltrata(rubrica);
      return;
    }

    const q = cerca.toLowerCase();
    setFiltrata(
      rubrica.filter((entry) => {
        const nome = entry.user?.displayName?.toLowerCase() ?? "";
        const mail = (entry.user?.mail ?? "").toLowerCase();
        const numero = entry.numeroInterno ?? "";
        const reparto = entry.user?.department?.toLowerCase() ?? "";
        return (
          nome.includes(q) ||
          mail.includes(q) ||
          numero.includes(q) ||
          reparto.includes(q)
        );
      })
    );
  }, [cerca, rubrica]);

  if (loading) return <div>Caricamento...</div>;
  if (errore) return <div>{errore}</div>;

  return (
    <div className={styles.rubrica}>
      <input
        className={styles.cerca}
        type="text"
        placeholder="Cerca per nome, reparto, numero..."
        value={cerca}
        onChange={(e) => setCerca(e.target.value)}
      />
      <div className={styles.griglia}>
        {filtrata.map((entry, i) => (
          <PersonCard key={i} entry={entry} />
        ))}
      </div>
    </div>
  );
};

// Merge tra utenti Graph e voci SP.
// - Parte dagli utenti Graph come fonte principale
// - Aggiunge le voci SP senza match (servizi) in fondo
const merge = (utenti: User[], voci: IVociLista[]): IRubricaUser[] => {
  // Mappa mail → numeroInterno per lookup veloce O(1)
  const mapSP = new Map<string, string>();
  voci.forEach((v) => mapSP.set(v.mail.toLowerCase(), v.numeroInterno));

  // Traccia le mail SP già usate nel match con Graph
  const mailUsate = new Set<string>();

  // 1. Tutti gli utenti Graph, arricchiti con il numero interno se presente
  const daGraph: IRubricaUser[] = utenti.map((u) => {
    const mailKey = u.mail?.toLowerCase() ?? "";
    const numeroInterno = mapSP.get(mailKey) ?? null;
    if (numeroInterno) mailUsate.add(mailKey);
    return { user: u, numeroInterno };
  });

  // 2. Voci SP senza match Graph (servizi tipo "Centralino")
  const servizi: IRubricaUser[] = voci
    .filter((v) => !mailUsate.has(v.mail.toLowerCase()))
    .map((v) => ({ user: null, numeroInterno: v.numeroInterno }));

  return [...daGraph, ...servizi];
};

export default Rubrica;