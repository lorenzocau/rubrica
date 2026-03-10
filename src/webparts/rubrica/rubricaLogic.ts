import { User } from "@microsoft/microsoft-graph-types";
import { IContattoInterno, IRubricaEntry } from "./types";


export function merge(datiGraph: User[], datiLista: IContattoInterno[]): IRubricaEntry[] {
    const rubrica: IRubricaEntry[] = [];
    const internimap = new Map();

    let entry: IRubricaEntry;
    for (const datoLista of datiLista)
        internimap.set(datoLista.contatto.toLowerCase(), datoLista.interno);

    for (const user of datiGraph) {
        if (internimap.get(user.mail?.toLowerCase()) != undefined) {
            entry = { tipo: "utente", user: user, interno: internimap.get(user.mail?.toLowerCase()) };
            internimap.delete(user.mail?.toLowerCase());
        }
        else {
            entry = { tipo: "utente", user: user, interno: null };
        }
        rubrica.push(entry);
    }

    for (const [contatto, interno] of Array.from(internimap)) {
        entry = { tipo: "servizio", nome: contatto, interno: interno };
        rubrica.push(entry);
    }

    return rubrica;
}

export function filtra(input: string, risultati: IRubricaEntry[]): IRubricaEntry[] {
    if (input.trim() === "") return []; // nulla mostrato finché non si cerca

    const q = input.toLowerCase();
    const filtrati: IRubricaEntry[] = [];

    for (const entry of risultati) {
        if (entry.tipo === "utente") {
            if (entry.user.displayName?.toLowerCase().includes(q) ||
                entry.user.mail?.toLowerCase().includes(q) ||
                entry.user.department?.toLowerCase().includes(q) ||
                entry.interno?.toLowerCase().includes(q) ||
                entry.user.jobTitle?.toLowerCase().includes(q))

                filtrati.push(entry);
        }
        else {
            if (entry.nome.toLowerCase().includes(q) ||
                entry.interno.toLowerCase().includes(q) ||
                "servizio".includes(q))

                filtrati.push(entry);
        }
    }
    return filtrati;
}