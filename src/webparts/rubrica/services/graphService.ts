import { getGraph } from "../pnpjsConfig";
import { User } from "@microsoft/microsoft-graph-types";

export async function getGraphUsers(): Promise<User[]> {
    const graph = getGraph();

    try {
        return await graph.users
            .select("mail", "id", "displayName", "department", "jobTitle", "mobilePhone", "businessPhones")
            .filter("accountEnabled eq true")
            .top(999)();
    } catch (e) {
        console.warn("Graph non raggiungibile. Gli utenti non saranno disponibili.", e);
        return [];
    }
}