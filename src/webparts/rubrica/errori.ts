// Trasforma gli errori tecnici di PnPjs/Graph/SharePoint in messaggi comprensibili.
// Sono funzioni pure: ricevono l'errore e restituiscono testo, così si possono testare.
// L'errore originale resta comunque nella console (vedi conAvviso in Rubrica.tsx).

export interface IDettaglioErrore {
    stato?: number;  // codice HTTP, se l'errore ne ha uno (401, 403, 404...)
    rete: boolean;   // true = la richiesta non è arrivata a destinazione
    testo: string;   // messaggio tecnico, accorciato
}

const MAX_TESTO = 160;

// Legge un errore "sconosciuto" senza fidarsi della sua forma: PnPjs lancia
// HttpRequestError (con .status), il browser lancia TypeError se manca la rete.
export function leggiErrore(e: unknown): IDettaglioErrore {
    const err = (e ?? {}) as { status?: unknown; response?: { status?: unknown }; message?: unknown };

    const messaggio = typeof err.message === "string" ? err.message.trim() : "";

    let stato: number | undefined;
    if (typeof err.status === "number") stato = err.status;
    else if (typeof err.response?.status === "number") stato = err.response.status;
    else {
        // ripiego: cerca un codice nel testo ("Error making HTTP request: 404")
        const m = /(?:http|status|request)\D{0,15}\b([1-5]\d{2})\b/i.exec(messaggio);
        if (m) stato = Number(m[1]);
    }

    const rete = stato === undefined && (
        e instanceof TypeError || /failed to fetch|networkerror|load failed|network request failed/i.test(messaggio)
    );

    const testo = messaggio.length > MAX_TESTO ? messaggio.slice(0, MAX_TESTO) + "…" : messaggio;
    return { stato, rete, testo };
}

function codice(d: IDettaglioErrore): string {
    return d.stato !== undefined ? ` (errore ${d.stato})` : "";
}

// Aggiunge il testo tecnico solo quando non si è riusciti a capire la causa.
function conDettaglio(base: string, d: IDettaglioErrore): string {
    return d.testo !== "" ? `${base} Dettaglio: ${d.testo}` : base;
}

// --- Microsoft Graph (elenco utenti) ---
export function erroreGraph(e: unknown): string {
    const d = leggiErrore(e);
    const effetto = "La rubrica mostra solo servizi ed esterni.";

    if (d.rete) {
        return `Impossibile raggiungere Microsoft Graph: controlla la connessione e ricarica la pagina. ${effetto}`;
    }
    if (d.stato === 401) {
        return `Sessione non valida con Microsoft Graph${codice(d)}: ricarica la pagina o riaccedi. ${effetto}`;
    }
    if (d.stato === 403) {
        return `Microsoft Graph ha negato l'accesso${codice(d)}: il permesso User.Read.All probabilmente non è stato approvato in SharePoint Admin Center (API access). ${effetto}`;
    }
    if (d.stato === 429) {
        return `Microsoft Graph sta limitando le richieste${codice(d)}: riprova tra qualche minuto. ${effetto}`;
    }
    if (d.stato !== undefined && d.stato >= 500) {
        return `Microsoft Graph non sta rispondendo${codice(d)}: riprova più tardi. ${effetto}`;
    }
    return conDettaglio(`Recupero utenti da Microsoft Graph non riuscito${codice(d)}. ${effetto}`, d);
}

// Graph ha risposto ma senza utenti: non è un errore tecnico, ma di solito è un problema.
export const AVVISO_GRAPH_VUOTO =
    "Microsoft Graph non ha restituito nessun utente: controlla il permesso User.Read.All e che esistano account attivi.";

// --- Liste SharePoint ---
// "conseguenza" dice cosa manca all'utente, es. "gli interni dei servizi non sono disponibili".
export function erroreLista(e: unknown, nomeLista: string, conseguenza: string): string {
    const d = leggiErrore(e);
    const lista = `"${nomeLista}"`;

    if (d.rete) {
        return `Impossibile raggiungere SharePoint per leggere la lista ${lista}: controlla la connessione e ricarica la pagina. ${conseguenza}`;
    }
    if (d.stato === 404) {
        return `Lista ${lista} non trovata${codice(d)}: controlla il nome nelle impostazioni (se è in un altro sito indica il percorso, es. /sites/intranet/Lists/Nome). ${conseguenza}`;
    }
    if (d.stato === 401 || d.stato === 403) {
        return `Non hai i permessi per leggere la lista ${lista}${codice(d)}: chiedi accesso in lettura a chi gestisce il sito. ${conseguenza}`;
    }
    // colonna mancante: SharePoint risponde 400 con "field or property ... does not exist"
    if (d.stato === 400 || /field or property|does not exist/i.test(d.testo)) {
        return `La lista ${lista} esiste ma non ha le colonne "Contatto" e "NumeroInterno"${codice(d)}: controlla i nomi interni delle colonne. ${conseguenza}`;
    }
    if (d.stato === 429) {
        return `SharePoint sta limitando le richieste${codice(d)}: riprova tra qualche minuto. ${conseguenza}`;
    }
    if (d.stato !== undefined && d.stato >= 500) {
        return `SharePoint non sta rispondendo${codice(d)}: riprova più tardi. ${conseguenza}`;
    }
    return conDettaglio(`Lettura della lista ${lista} non riuscita${codice(d)}. ${conseguenza}`, d);
}

// Errore non previsto in tutto il caricamento.
export function erroreImprevisto(e: unknown): string {
    const d = leggiErrore(e);
    return conDettaglio(`Errore imprevisto nel caricamento della rubrica${codice(d)}.`, d);
}