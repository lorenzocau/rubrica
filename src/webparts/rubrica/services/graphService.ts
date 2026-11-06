import { getGraph } from "../pnpjsConfig";
import { User } from "@microsoft/microsoft-graph-types";
import { leggiCache, scriviCache } from "./cache";

// la versione nella chiave scarta le vecchie voci quando cambiano i campi letti
const CACHE_CHIAVE = "graphUsers:v2";
const CACHE_TTL_MS = 10 * 60 * 1000; // 10 minuti

// Se Graph non è raggiungibile (o i permessi non sono approvati) la chiamata
// lancia un errore: è il chiamante (Rubrica.tsx) a mostrare l'avviso.
export async function getGraphUsers(): Promise<User[]> {
    const incache = leggiCache<User[]>(CACHE_CHIAVE, CACHE_TTL_MS);
    if (incache) return incache;

    const graph = getGraph();

    const utenti: User[] = await graph.users
        .select("mail", "id", "displayName", "officeLocation", "mobilePhone", "businessPhones", "userType")
        .filter("accountEnabled eq true")
        .top(999)();

    // una risposta vuota non si salva: meglio riprovare alla prossima visita
    if (utenti.length > 0) scriviCache(CACHE_CHIAVE, utenti);

    return utenti;
}
