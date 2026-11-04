import { User } from "@microsoft/microsoft-graph-types";
import { IContattoInterno, IRubricaEntry } from "./types";

const SEPARATORE_INTERNI = " / ";

// Voce della lista già normalizzata: il nome mantiene le maiuscole originali,
// la chiave di ricerca (minuscola) vive solo nella Map.
interface IVoce {
    nome: string;
    interni: string[];
}

export function merge(users: User[], vociLista: IContattoInterno[]): IRubricaEntry[] {
    // chiave = contatto in minuscolo e senza spazi ai bordi
    const voci = new Map<string, IVoce>();

    for (const voceLista of vociLista) {
        const nome = (voceLista.contatto ?? "").trim();
        const interno = (voceLista.interno ?? "").trim();

        // righe incomplete: ignorate
        if (nome === "" || interno === "") continue;

        const chiave = nome.toLowerCase();
        const esistente = voci.get(chiave);

        if (!esistente) {
            voci.set(chiave, { nome, interni: [interno] });
        }
        else if (esistente.interni.indexOf(interno) === -1) {
            // stesso contatto con più interni: li teniamo tutti
            esistente.interni.push(interno);
        }
    }

    const rubrica: IRubricaEntry[] = [];
    const usate = new Set<string>();

    for (const user of users) {
        const mail = user.mail?.trim().toLowerCase();
        let interno: string | null = null;

        if (mail) {
            const voce = voci.get(mail);
            if (voce) {
                interno = voce.interni.join(SEPARATORE_INTERNI);
                usate.add(mail);
            }
        }

        rubrica.push({ tipo: "utente", user: user, interno: interno });
    }

    // voci senza utente corrispondente in Graph: servizi
    // (se Graph fallisce, finiscono qui anche le voci degli utenti)
    voci.forEach((voce, chiave) => {
        if (!usate.has(chiave)) {
            rubrica.push({
                tipo: "servizio",
                nome: voce.nome,
                interno: voce.interni.join(SEPARATORE_INTERNI),
            });
        }
    });

    return rubrica;
}

// minuscolo e senza accenti: "Qualità" e "qualita" diventano uguali
function normalizza(s?: string | null): string {
    return (s ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase();
}

export function filtra(input: string, risultati: IRubricaEntry[]): IRubricaEntry[] {
    // più parole: ogni parola deve comparire in almeno un campo, in qualsiasi ordine
    const parole = normalizza(input).split(/\s+/).filter(p => p !== "");
    if (parole.length === 0) return []; // nulla mostrato finché non si cerca

    return risultati.filter(entry => {
        const testo = entry.tipo === "utente"
            ? [entry.user.displayName, entry.user.mail, entry.user.department, entry.user.jobTitle, entry.interno]
            : [entry.nome, entry.interno];

        const haystack = testo.map(t => normalizza(t)).join(" ");

        return parole.every(p =>
            haystack.indexOf(p) !== -1 ||
            // "servizio"/"servizi" trovano tutti i servizi, ma solo come parola intera
            (entry.tipo === "servizio" && (p === "servizio" || p === "servizi"))
        );
    });
}