import { getGraph } from "../pnpjsConfig";
import { User } from "@microsoft/microsoft-graph-types";

export const getGraphUsers = async (): Promise<User[]> => {
    const graph = getGraph();

    return await graph.users
        .select(
            "mail", // key
            "id",
            "displayName",
            "department",
            "jobTitle",
            "mobilePhone",
            "businessPhones"
        )
        .filter("accountEnabled eq true")
        .top(999)();
}