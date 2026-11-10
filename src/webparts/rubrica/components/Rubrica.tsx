import * as React from "react";
import { useState, useEffect, useMemo } from "react";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { Text } from "@fluentui/react/lib/Text";
import { Checkbox } from "@fluentui/react/lib/Checkbox";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";

// dati reali (i dati mock vengono caricati a richiesta, vedi getSorgenti)
import { getGraphUsers } from "../services/graphService";
import { getVociEsterni, getVociLista } from "../services/spService";
import { IContattoEsterno, IContattoInterno, IRubricaEntry, IRubricaWebPartProps, UtenteGraph } from "../types";
import PersonCard from "./PersonCard";
import { applicaVista, filtra, merge, normalizzaPageSize, numeroBreveNonAssegnato, risolviNomeLista } from "../rubricaLogic";
import { PROPRIETA_DEFAULT } from "../proprieta";
import { AVVISO_GRAPH_VUOTO, COLONNE_LISTA_DEFAULT, erroreGraph, erroreImprevisto, erroreLista } from "../errori";

// colonne attese nelle liste, per il messaggio d'errore sulle colonne mancanti
const COLONNE_SERVIZI = [...COLONNE_LISTA_DEFAULT, "Reparto"];
const COLONNE_ESTERNI = ["Nome", "Numero", "Numero breve"];

interface IEsito<T> {
  dati: T[];
  ok: boolean;
}

// Esegue una sorgente dati: se fallisce registra l'avviso e restituisce [],
// così le altre sorgenti vengono comunque mostrate. "ok" dice se è riuscita.
// "descrivi" trasforma l'errore tecnico in un messaggio per l'utente (vedi errori.ts);
// l'errore originale resta nella console per chi deve indagare.
async function conAvviso<T>(
  promessa: Promise<T[]>,
  descrivi: (errore: unknown) => string,
  avvisi: string[]
): Promise<IEsito<T>> {
  try {
    return { dati: await promessa, ok: true };
  } catch (e) {
    const avviso = descrivi(e);
    console.warn("[Rubrica]", avviso, e);
    avvisi.push(avviso);
    return { dati: [], ok: false };
  }
}

interface ISorgenti {
  utenti: () => Promise<UtenteGraph[]>;
  servizi: (lista: string) => Promise<IContattoInterno[]>;
  esterni: (lista: string) => Promise<IContattoEsterno[]>;
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
  return { utenti: getGraphUsers, servizi: getVociLista, esterni: getVociEsterni };
}

function Rubrica(props: IRubricaWebPartProps) {
  const [datiGraph, setDatiGraph] = useState<UtenteGraph[]>([]);
  const [datiServizi, setDatiServizi] = useState<IContattoInterno[]>([]);
  const [datiEsterni, setDatiEsterni] = useState<IContattoEsterno[]>([]);
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
            erroreGraph,
            avvisiCaricamento
          ),
          conAvviso(
            sorgenti.servizi(listaServizi),
            e => erroreLista(e, listaServizi, "I servizi non sono disponibili.", COLONNE_SERVIZI),
            avvisiCaricamento
          ),
          props.mostraEsterni
            ? conAvviso(
              sorgenti.esterni(listaEsterni),
              e => erroreLista(e, listaEsterni, "I contatti esterni non sono disponibili.", COLONNE_ESTERNI),
              avvisiCaricamento
            )
            : Promise.resolve<IEsito<IContattoEsterno>>({ dati: [], ok: false }),
        ]);

        if (annullato) return;
        // Graph ha risposto ma senza utenti: quasi sempre è un problema di permessi
        if (graph.ok && graph.dati.length === 0 && !props.usaDatiMock) {
          avvisiCaricamento.push(AVVISO_GRAPH_VUOTO);
        }
        setDatiGraph(graph.dati);
        setDatiServizi(servizi.dati);
        setDatiEsterni(esterni.dati);
        setEsterniOk(esterni.ok);
      } catch (e) {
        console.error("[Rubrica] Errore imprevisto nel caricamento.", e);
        avvisiCaricamento.push(erroreImprevisto(e));
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
      nascondiSenzaTelefono: props.nascondiSenzaTelefono,
      includiEsterni: esterniDisponibili && includiEsterni,
    }),
    [rubricaCompleta, props.soloMembri, props.nascondiSenzaMail, props.nascondiSenzaTelefono, esterniDisponibili, includiEsterni]
  );

  const ricercaPulita = ricerca.trim();
  const mostraTutto = props.mostraElencoCompleto && ricercaPulita === "";

  const filtrati = useMemo<IRubricaEntry[]>(
    () => (mostraTutto ? visibile : filtra(ricerca, visibile)),
    [mostraTutto, ricerca, visibile]
  );

  // Cercando un numero breve dentro l'intervallo di quelli degli esterni ma non assegnato
  // a nessuno, si avvisa. Si guarda l'elenco completo (non quello filtrato) e solo se gli
  // esterni sono inclusi: con la checkbox spenta potrebbe essere proprio il numero di un esterno.
  const numeroLibero = useMemo(
    () => (esterniDisponibili && includiEsterni && ricercaPulita !== "")
      ? numeroBreveNonAssegnato(ricercaPulita, rubricaCompleta)
      : null,
    [esterniDisponibili, includiEsterni, ricercaPulita, rubricaCompleta]
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
      {numeroLibero && (
        <MessageBar messageBarType={MessageBarType.info}>
          Il numero breve <strong>{numeroLibero.numero}</strong> rientra nell’intervallo dei numeri
          brevi degli esterni ({numeroLibero.min}–{numeroLibero.max}) ma non è associato a nessun contatto.
        </MessageBar>
      )}
      {filtrati.length === 0 && !numeroLibero && (ricercaPulita !== "" || mostraTutto) && (
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