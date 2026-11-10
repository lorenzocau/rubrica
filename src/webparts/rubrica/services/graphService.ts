import { getGraph } from "../pnpjsConfig";
import { UtenteGraph } from "../types";
import { leggiCache, scriviCache } from "./cache";

// la versione nella chiave scarta le vecchie voci quando cambiano i campi letti
// (v4: reparto (department) + jobTitle al posto di officeLocation, aggiunto faxNumber)
const CACHE_CHIAVE = "graphUsers:v4";
const CACHE_TTL_MS = 10 * 60 * 1000; // 10 minuti

// Se Graph non è raggiungibile (o i permessi non sono approvati) la chiamata
// lancia un errore: è il chiamante (Rubrica.tsx) a mostrare l'avviso.
//
// Campi telefonici in AD:
//   - businessPhones[0] = "Telephone number" -> da cui si ricava l'interno (ultime 3 cifre)
//   - mobilePhone       = cellulare aziendale
//   - faxNumber         = numero breve del cellulare aziendale (non è tra i campi restituiti
//                         di default da Graph: per questo va richiesto nel select)
export async function getGraphUsers(): Promise<UtenteGraph[]> {
    const incache = leggiCache<UtenteGraph[]>(CACHE_CHIAVE, CACHE_TTL_MS);
    if (incache) return incache;

    const graph = getGraph();

    const utenti: UtenteGraph[] = await graph.users
        .select("mail", "id", "displayName", "department", "jobTitle", "mobilePhone", "faxNumber", "businessPhones", "userType")
        .filter("accountEnabled eq true")
        .top(999)();

    // una risposta vuota non si salva: meglio riprovare alla prossima visita
    if (utenti.length > 0) scriviCache(CACHE_CHIAVE, utenti);

    return utenti;
}
