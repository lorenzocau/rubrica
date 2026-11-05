import * as React from "react";
import { useState, useEffect, useMemo } from "react";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { Text } from "@fluentui/react/lib/Text";
import { Checkbox } from "@fluentui/react/lib/Checkbox";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { User } from "@microsoft/microsoft-graph-types";

// dati reali (i dati mock vengono caricati a richiesta, vedi getSorgenti)
import { getGraphUsers } from "../services/graphService";
import { getVociLista } from "../services/spService";
import { IContattoInterno, IRubricaEntry, IRubricaWebPartProps } from "../types";
import PersonCard from "./PersonCard";
import { applicaVista, filtra, merge, normalizzaPageSize, risolviNomeLista } from "../rubricaLogic";
import { PROPRIETA_DEFAULT } from "../proprieta";

interface IEsito<T> {
  dati: T[];
  ok: boolean;
}

// Esegue una sorgente dati: se fallisce registra l'avviso e restituisce [],
// così le altre sorgenti vengono comunque mostrate. "ok" dice se è riuscita.
async function conAvviso<T>(promessa: Promise<T[]>, avviso: string, avvisi: string[]): Promise<IEsito<T>> {
  try {
    return { dati: await promessa, ok: true };
  } catch (e) {
    console.warn(avviso, e);
    avvisi.push(avviso);
    return { dati: [], ok: false };
  }
}

interface ISorgenti {
  utenti: () => Promise<User[]>;
  servizi: (lista: string) => Promise<IContattoInterno[]>;
  esterni: (lista: string) => Promise<IContattoInterno[]>;
}

// I dati finti stanno in un chunk separato: il browser lo scarica solo se
// "usaDatiMock" è attivo, quindi in condizioni normali non pesano sulla pagina.
async function getSorgenti(usaMock: boolean): Promise<ISorgenti> {
  if (usaMock) {
    const mock = await import(/* webpackChunkName: "rubrica-mockdata" */ "../mockdata");
    return {
      utenti: mock.getMockGraphUsers,
      servizi: () => mock.getMockVociLista(),
      esterni: () => mock.getMockEsterni(),
    };
  }
  return { utenti: getGraphUsers, servizi: getVociLista, esterni: getVociLista };
}

function Rubrica(props: IRubricaWebPartProps) {
  const [datiGraph, setDatiGraph] = useState<User[]>([]);
  const [datiServizi, setDatiServizi] = useState<IContattoInterno[]>([]);
  const [datiEsterni, setDatiEsterni] = useState<IContattoInterno[]>([]);
  const [esterniOk, setEsterniOk] = useState<boolean>(false);
  const [includiEsterni, setIncludiEsterni] = useState<boolean>(props.esterniInizialmenteVisibili);
  const [ricerca, setRicerca] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(true);
  const [avvisi, setAvvisi] = useState<string[]>([]);
  const [pagina, setPagina] = useState<number>(0);

  const listaServizi = risolviNomeLista(props.listaServizi, PROPRIETA_DEFAULT.listaServizi);
  const listaEsterni = risolviNomeLista(props.listaEsterni, PROPRIETA_DEFAULT.listaEsterni);
  const pageSize = normalizzaPageSize(props.risultatiPerPagina);

  // Ricarica i dati solo se cambia qualcosa che riguarda DA DOVE si leggono.
  // I filtri sugli utenti e le opzioni di visualizzazione non richiedono nuove chiamate.
  useEffect(() => {
    let annullato = false; // evita di applicare una risposta vecchia se nel frattempo le impostazioni sono cambiate
    const avvisiCaricamento: string[] = [];

    async function caricaDati() {
      setLoading(true);
      try {
        const sorgenti = await getSorgenti(props.usaDatiMock);

        const [graph, servizi, esterni] = await Promise.all([
          conAvviso(
            sorgenti.utenti(),
            "Recupero utenti da Microsoft Graph non riuscito: la rubrica potrebbe essere incompleta.",
            avvisiCaricamento
          ),
          conAvviso(
            sorgenti.servizi(listaServizi),
            `Lista "${listaServizi}" non trovata o non accessibile in questo sito: gli interni non sono disponibili.`,
            avvisiCaricamento
          ),
          props.mostraEsterni
            ? conAvviso(
                sorgenti.esterni(listaEsterni),
                `Lista "${listaEsterni}" non trovata o non accessibile in questo sito: i contatti esterni non sono disponibili.`,
                avvisiCaricamento
              )
            : Promise.resolve<IEsito<IContattoInterno>>({ dati: [], ok: false }),
        ]);

        if (annullato) return;
        setDatiGraph(graph.dati);
        setDatiServizi(servizi.dati);
        setDatiEsterni(esterni.dati);
        setEsterniOk(esterni.ok);
      } catch (e) {
        console.error("Errore imprevisto nel caricamento della rubrica.", e);
        avvisiCaricamento.push("Errore imprevisto nel caricamento della rubrica.");
      } finally {
        if (!annullato) {
          setAvvisi(avvisiCaricamento);
          setLoading(false);
        }
      }
    }

    void caricaDati();
    return () => { annullato = true; };
  }, [props.usaDatiMock, props.mostraEsterni, listaServizi, listaEsterni]);

  // se dal pannello cambia lo stato iniziale della checkbox, la si riallinea
  useEffect(() => {
    setIncludiEsterni(props.esterniInizialmenteVisibili);
  }, [props.esterniInizialmenteVisibili]);

  // la checkbox esiste solo se la funzione è attiva E la lista Esterni è stata letta
  const esterniDisponibili = props.mostraEsterni && esterniOk;

  const rubricaCompleta = useMemo<IRubricaEntry[]>(
    () => merge(datiGraph, datiServizi, datiEsterni),
    [datiGraph, datiServizi, datiEsterni]
  );

  const visibile = useMemo<IRubricaEntry[]>(
    () => applicaVista(rubricaCompleta, {
      soloMembri: props.soloMembri,
      nascondiSenzaMail: props.nascondiSenzaMail,
      includiEsterni: esterniDisponibili && includiEsterni,
    }),
    [rubricaCompleta, props.soloMembri, props.nascondiSenzaMail, esterniDisponibili, includiEsterni]
  );

  const ricercaPulita = ricerca.trim();
  const mostraTutto = props.mostraElencoCompleto && ricercaPulita === "";

  const filtrati = useMemo<IRubricaEntry[]>(
    () => (mostraTutto ? visibile : filtra(ricerca, visibile)),
    [mostraTutto, ricerca, visibile]
  );

  useEffect(() => {
    setPagina(0); // reset alla prima pagina ad ogni cambio di ricerca, filtri o dimensione pagina
  }, [ricerca, visibile, pageSize]);

  if (loading) return <Spinner size={SpinnerSize.medium} label="Caricamento rubrica..." />;

  const utenti = visibile.filter(e => e.tipo === "utente").length;
  const servizi = visibile.filter(e => e.tipo === "servizio").length;
  const esterni = visibile.filter(e => e.tipo === "esterno").length;

  const totPagine = Math.ceil(filtrati.length / pageSize);
  const paginaCorr = Math.min(pagina, Math.max(totPagine - 1, 0));
  const inizio = paginaCorr * pageSize;
  const paginaVisibile = filtrati.slice(inizio, inizio + pageSize);

  return (
    <>
      {props.usaDatiMock && (
        <MessageBar messageBarType={MessageBarType.severeWarning}>
          Modalità dati di prova attiva: i contatti mostrati non sono reali.
        </MessageBar>
      )}

      {avvisi.map((avviso, i) => (
        <MessageBar key={i} messageBarType={MessageBarType.warning}>{avviso}</MessageBar>
      ))}

      {props.titolo.trim() !== "" && (
        <Text as="h2" variant="xLarge" block style={{ marginBottom: 8 }}>{props.titolo}</Text>
      )}

      {/* ── Contatore caricati ── */}
      {props.mostraContatore && (
        <Text variant="small" style={{ color: "#605e5c", display: "block", marginBottom: 8 }}>
          Rubrica: <strong>{utenti}</strong> utenti · <strong>{servizi}</strong> servizi
          {esterniDisponibili && includiEsterni && <> · <strong>{esterni}</strong> esterni</>}
        </Text>
      )}

      <SearchBox
        placeholder={props.placeholderRicerca}
        ariaLabel="Cerca nella rubrica"
        value={ricerca}
        onChange={(_e, text) => setRicerca(text ?? "")}
        onClear={() => setRicerca("")}
      />

      {esterniDisponibili && (
        <Checkbox
          label="Includi contatti esterni"
          checked={includiEsterni}
          onChange={(_e, checked) => setIncludiEsterni(!!checked)}
          styles={{ root: { margin: "8px 0" } }}
        />
      )}

      {/* ── Contatore risultati + pagina ── */}
      {(ricercaPulita !== "" || mostraTutto) && filtrati.length > 0 && (
        <Text variant="small" style={{ color: "#605e5c", display: "block", margin: "6px 0" }}>
          {filtrati.length} risultat{filtrati.length === 1 ? "o" : "i"}
          {totPagine > 1 && ` — pagina ${paginaCorr + 1} di ${totPagine}`}
        </Text>
      )}

      {/* ── Messaggi di stato ── */}
      {!mostraTutto && ricercaPulita === "" && (
        <Text variant="small" style={{ color: "#605e5c", display: "block", margin: "12px 0" }}>
          Scrivi nel campo di ricerca per trovare un contatto.
        </Text>
      )}
      {filtrati.length === 0 && (ricercaPulita !== "" || mostraTutto) && (
        <Text variant="medium" style={{ color: "#605e5c", display: "block", margin: "12px 0" }}>
          {mostraTutto ? "Nessun contatto disponibile." : `Nessun risultato per “${ricercaPulita}”.`}
        </Text>
      )}

      {paginaVisibile.map((entry, i) => <PersonCard key={i} entry={entry} opzioni={props} />)}

      {/* ── Paginazione ── */}
      {totPagine > 1 && (
        <div style={{ display: "flex", gap: 8, padding: "12px 0", justifyContent: "center" }}>
          <DefaultButton
            disabled={paginaCorr === 0}
            onClick={() => setPagina(paginaCorr - 1)}
          >← Precedente</DefaultButton>
          <DefaultButton
            disabled={paginaCorr >= totPagine - 1}
            onClick={() => setPagina(paginaCorr + 1)}
          >Successiva →</DefaultButton>
        </div>
      )}
    </>
  );
}

export default Rubrica;
