// WebPartContext è il contesto SPFx: contiene informazioni sull'utente loggato,
// il sito corrente, e i token di autenticazione. È il "passaporto" che permette
// a PnPjs di fare chiamate autenticate a SharePoint e Graph senza che noi
// dobbiamo gestire manualmente token o header.
import { WebPartContext } from "@microsoft/sp-webpart-base";

// --- SharePoint ---
// spfi    = SharePoint Factory Interface, il punto di ingresso per le chiamate SP
// SPFI    = il tipo TypeScript dell'istanza creata da spfi()
// spSPFx  = il "behavior" che collega PnPjs al contesto SPFx per SP.
//           Alias necessario perché sia @pnp/sp che @pnp/graph esportano
//           qualcosa chiamato "SPFx" — senza alias TypeScript andrebbe in conflitto.
import { spfi, SPFI, SPFx as spSPFx } from "@pnp/sp";

// PnPjs usa "selective imports": importi solo i moduli che ti servono.
// Questo riduce il bundle finale (tree shaking).
// Senza questi import, le chiamate tipo sp.web.lists.(...).items non esisterebbero.
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

// --- Graph ---
// graphfi  = Graph Factory Interface, il punto di ingresso per le chiamate Graph
// GraphFI  = il tipo TypeScript dell'istanza creata da graphfi()
// graphSPFx = il "behavior" che collega PnPjs al contesto SPFx per Graph.
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";

// Modulo selettivo per gli utenti — necessario per graph.users()
import "@pnp/graph/users";

// --- Logging ---
// PnPLogging è un behavior opzionale che stampa le chiamate API in console.
// LogLevel.Warning significa: logga solo Warning ed Error, non tutto.
import { LogLevel, PnPLogging } from "@pnp/logging";

// Variabili globali che contengono le istanze configurate.
// Sono "null" finché non viene chiamato getSP/getGraph con il context.
// Questo è un pattern Singleton: l'istanza viene creata una sola volta
// (in onInit del WebPart) e poi riutilizzata da tutti i service senza
// dover ri-passare il context ogni volta.
let _sp: SPFI | null = null;
let _graph: GraphFI | null = null;

// getSP va chiamata UNA VOLTA con il context (da RubricaWebPart.ts → onInit).
// Tutte le chiamate successive (dai service) la chiamano senza argomenti
// e ricevono l'istanza già configurata.
export const getSP = (context?: WebPartContext): SPFI => {
    if (context != null) {
        // spfi()          → crea una nuova istanza
        // .using(spSPFx)  → la configura con il context SPFx (autenticazione, URL del sito)
        // .using(PnPLogging) → aggiunge il logging in console
        _sp = spfi().using(spSPFx(context)).using(PnPLogging(LogLevel.Warning));
    }
    return _sp!;
};

// Stesso pattern di getSP, ma per Microsoft Graph.
export const getGraph = (context?: WebPartContext): GraphFI => {
    if (context != null) {
        _graph = graphfi().using(graphSPFx(context)).using(PnPLogging(LogLevel.Warning));
    }
    return _graph!;
};