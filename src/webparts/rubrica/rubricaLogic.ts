import { User } from "@microsoft/microsoft-graph-types";
import { IContattoInterno, IRubricaEntry } from "./types";

const SEPARATORE_INTERNI = " / ";

export const PAGE_SIZE_DEFAULT = 30;
export const PAGE_SIZE_MIN = 5;
export const PAGE_SIZE_MAX = 100;

// Voce della lista già normalizzata: il nome mantiene le maiuscole originali,
// la chiave di ricerca (minuscola) vive solo nella Map.
interface IVoce {
    nome: string;
    interni: string[];
}

// Normalizza le righe di una lista: ignora quelle incomplete, unisce i duplicati
// (case-insensitive) e tiene tutti gli interni distinti.
function raggruppa(vociLista: IContattoInterno[]): Map<string, IVoce> {
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

    return voci;
}

// Ordine alfabetico senza distinguere maiuscole e accenti ("Álvaro" prima di "Bruno").
const collator = new Intl.Collator("it", { sensitivity: "base" });

function nomeOrdinamento(entry: IRubricaEntry): string {
    return entry.tipo === "utente"
        ? (entry.user.displayName ?? entry.user.mail ?? "")
        : entry.nome;
}

export function ordina(entries: IRubricaEntry[]): IRubricaEntry[] {
    // l'indice come ultimo criterio rende l'ordinamento stabile a parità di nome
    return entries
        .map((entry, indice) => ({ entry, indice }))
        .sort((a, b) =>
            collator.compare(nomeOrdinamento(a.entry), nomeOrdinamento(b.entry)) || a.indice - b.indice
        )
        .map(x => x.entry);
}

export function merge(
    users: User[],
    vociLista: IContattoInterno[],
    vociEsterni: IContattoInterno[] = []
): IRubricaEntry[] {
    const voci = raggruppa(vociLista);

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

    // esterni: lista separata, mai abbinata agli utenti Graph
    raggruppa(vociEsterni).forEach(voce => {
        rubrica.push({
            tipo: "esterno",
            nome: voce.nome,
            interno: voce.interni.join(SEPARATORE_INTERNI),
        });
    });

    return ordina(rubrica);
}

export interface IOpzioniVista {
    soloMembri: boolean;        // nasconde gli ospiti (userType Guest)
    nascondiSenzaMail: boolean; // nasconde gli account senza mail
    includiEsterni: boolean;    // checkbox "Esterni"
}

// Applica i filtri di visualizzazione DOPO il merge: così un utente nascosto
// (es. un ospite) non fa ricomparire la sua voce di lista come "servizio".
export function applicaVista(rubrica: IRubricaEntry[], opz: IOpzioniVista): IRubricaEntry[] {
    return rubrica.filter(entry => {
        if (entry.tipo === "esterno") return opz.includiEsterni;

        if (entry.tipo === "utente") {
            if (opz.soloMembri && entry.user.userType === "Guest") return false;
            if (opz.nascondiSenzaMail && (entry.user.mail ?? "").trim() === "") return false;
        }

        return true;
    });
}

// minuscolo e senza accenti: "Qualità" e "qualita" diventano uguali
function normalizza(s?: string | null): string {
    return (s ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase();
}

// parole che selezionano un intero tipo di voce (solo come parola intera)
const PAROLE_TIPO: { [tipo: string]: string[] } = {
    servizio: ["servizio", "servizi"],
    esterno: ["esterno", "esterni"],
};

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
            // "servizio"/"servizi" e "esterno"/"esterni" trovano tutte le voci di quel tipo
            (entry.tipo !== "utente" && PAROLE_TIPO[entry.tipo].indexOf(p) !== -1)
        );
    });
}

// --- helper per le impostazioni del pannello ---

// Campo vuoto = usa il default.
export function risolviNomeLista(valore: string | undefined | null, predefinito: string): string {
    const v = (valore ?? "").trim();
    return v === "" ? predefinito : v;
}

// Messaggio d'errore per il pannello proprietà ("" = valido).
export function validaNomeLista(valore: string | undefined | null): string {
    const v = (valore ?? "").trim();
    if (v.length > 255) return "Massimo 255 caratteri.";
    if (/[<>"|*?]/.test(v)) return "Il nome contiene caratteri non validi.";
    return "";
}

// Tiene il numero di risultati per pagina dentro limiti sensati.
export function normalizzaPageSize(valore: number | undefined | null): number {
    if (typeof valore !== "number" || isNaN(valore)) return PAGE_SIZE_DEFAULT;
    return Math.min(PAGE_SIZE_MAX, Math.max(PAGE_SIZE_MIN, Math.round(valore)));
}
