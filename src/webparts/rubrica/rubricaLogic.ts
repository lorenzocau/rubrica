import { User } from "@microsoft/microsoft-graph-types";
import { IContattoInterno, IRubricaEntry } from "./types";


export function merge(users: User[], vociLista: IContattoInterno[]): IRubricaEntry[] {
    const rubrica: IRubricaEntry[] = [];
    const contattiMap = new Map();

    let entry: IRubricaEntry;
    for (const voceLista of vociLista)
        contattiMap.set(voceLista.contatto.toLowerCase(), voceLista.interno);

    for (const user of users) {
        const mail = user.mail?.toLowerCase();

        if (contattiMap.get(mail) != undefined) {
            entry = { tipo: "utente", user: user, interno: contattiMap.get(mail) };
            contattiMap.delete(mail);
        }
        else 
            entry = { tipo: "utente", user: user, interno: null };
        
        rubrica.push(entry);
    }
    // se graph non fa tutti vengono visti come servizio!
    for (const [contatto, interno] of Array.from(contattiMap)) {
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