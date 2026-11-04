import { getGraph } from "../pnpjsConfig";
import { User } from "@microsoft/microsoft-graph-types";

// Se Graph non è raggiungibile (o i permessi non sono approvati) la chiamata
// lancia un errore: è il chiamante (Rubrica.tsx) a mostrare l'avviso.
export async function getGraphUsers(): Promise<User[]> {
    const graph = getGraph();

    return await graph.users
        .select("mail", "id", "displayName", "department", "jobTitle", "mobilePhone", "businessPhones")
        .filter("accountEnabled eq true")
        .top(999)();
}