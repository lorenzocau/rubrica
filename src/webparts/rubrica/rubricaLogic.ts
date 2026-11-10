import { IContattoEsterno, IContattoInterno, IRubricaEntry, IRubricaUtente, UtenteGraph } from "./types";

export const SEPARATORE_INTERNI = " / ";

export const PAGE_SIZE_DEFAULT = 30;
export const PAGE_SIZE_MIN = 5;
export const PAGE_SIZE_MAX = 100;

// Un numero scritto per intero (cellulare, numero esterno) si cerca solo se chi cerca
// ha digitato almeno queste cifre: con meno, "101" troverebbe chiunque abbia "101"
// nel mezzo del cellulare, e i numeri brevi non si potrebbero più cercare con precisione.
export const MIN_CIFRE_NUMERO_LUNGO = 4;

// Voce della lista già normalizzata: il nome mantiene le maiuscole originali,
// la chiave di ricerca (minuscola) vive solo nella Map.
interface IVoce {
    nome: string;
    interni: string[];
    reparto: string; // primo reparto non vuoto trovato tra le righe duplicate
}

// Normalizza le righe di una lista: ignora quelle incomplete, unisce i duplicati
// (case-insensitive) e tiene tutti gli interni distinti.
function raggruppa(vociLista: IContattoInterno[]): Map<string, IVoce> {
    // chiave = contatto in minuscolo e senza spazi ai bordi
    const voci = new Map<string, IVoce>();

    for (const voceLista of vociLista) {
        const nome = (voceLista.contatto ?? "").trim();
        const interno = (voceLista.interno ?? "").trim();
        const reparto = (voceLista.reparto ?? "").trim();

        // righe incomplete: ignorate
        if (nome === "" || interno === "") continue;

        const chiave = nome.toLowerCase();
        const esistente = voci.get(chiave);

        if (!esistente) {
            voci.set(chiave, { nome, interni: [interno], reparto });
        }
        else {
            if (esistente.interni.indexOf(interno) === -1) {
                // stesso contatto con più interni: li teniamo tutti
                esistente.interni.push(interno);
            }
            if (esistente.reparto === "") esistente.reparto = reparto;
        }
    }

    return voci;
}

interface IVoceEsterna {
    nome: string;
    numeri: string[];
    numeriBrevi: string[];
}

function aggiungiDistinto(elenco: string[], valore: string): void {
    if (valore !== "" && elenco.indexOf(valore) === -1) elenco.push(valore);
}

// Come raggruppa, per la lista Esterni: una riga è valida con il nome e almeno un numero.
// Lo stesso nome (case-insensitive) su più righe diventa una sola voce con tutti i numeri.
function raggruppaEsterni(vociEsterni: IContattoEsterno[]): Map<string, IVoceEsterna> {
    const voci = new Map<string, IVoceEsterna>();

    for (const voce of vociEsterni) {
        const nome = (voce.nome ?? "").trim();
        const numero = (voce.numero ?? "").trim();
        const breve = (voce.numeroBreve ?? "").trim();

        if (nome === "" || (numero === "" && breve === "")) continue;

        const chiave = nome.toLowerCase();
        let esistente = voci.get(chiave);
        if (!esistente) {
            esistente = { nome, numeri: [], numeriBrevi: [] };
            voci.set(chiave, esistente);
        }
        aggiungiDistinto(esistente.numeri, numero);
        aggiungiDistinto(esistente.numeriBrevi, breve);
    }

    return voci;
}

// "101 / 102" -> ["101", "102"]. Serve per mostrare e confrontare i singoli numeri.
export function dividiNumeri(testo?: string | null): string[] {
    return (testo ?? "")
        .split(SEPARATORE_INTERNI)
        .map(n => n.trim())
        .filter(n => n !== "");
}

// Numero breve = ultime 3 cifre del "Telephone number" dell'utente (businessPhones[0] in Graph).
// Si guardano solo le cifre, quindi "+39 051 123 456" e "0511234 56" danno entrambi "456".
// Con meno di 3 cifre non c'è un numero breve affidabile: restituisce null.
export function estraiNumeroBreve(telefono?: string | null): string | null {
    const cifre = (telefono ?? "").replace(/\D/g, "");
    return cifre.length < 3 ? null : cifre.slice(-3);
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
    users: UtenteGraph[],
    vociLista: IContattoInterno[],
    vociEsterni: IContattoEsterno[] = []
): IRubricaEntry[] {
    const voci = raggruppa(vociLista);

    const rubrica: IRubricaEntry[] = [];
    const usate = new Set<string>();

    for (const user of users) {
        const mail = user.mail?.trim().toLowerCase();
        // di base l'interno è il numero breve preso dal telefono dell'utente;
        // una voce nella lista con la sua mail, se c'è, lo sostituisce
        let interno: string | null = estraiNumeroBreve(user.businessPhones?.[0]);

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
                reparto: voce.reparto,
            });
        }
    });

    // esterni: lista separata, mai abbinata agli utenti Graph
    raggruppaEsterni(vociEsterni).forEach(voce => {
        rubrica.push({
            tipo: "esterno",
            nome: voce.nome,
            numero: voce.numeri.join(SEPARATORE_INTERNI),
            numeroBreve: voce.numeriBrevi.join(SEPARATORE_INTERNI),
        });
    });

    return ordina(rubrica);
}

// Riga sotto il nome: "Reparto - Job title". Se ne manca uno si mostra solo l'altro.
export function descrizioneUtente(user: UtenteGraph): string {
    return [user.department, user.jobTitle]
        .map(t => (t ?? "").trim())
        .filter(t => t !== "")
        .join(" - ");
}

// Un utente ha "un numero di telefono" se ha almeno uno tra: interno (dalla lista o dal
// telefono fisso), cellulare aziendale, numero breve del cellulare.
export function haTelefono(entry: IRubricaUtente): boolean {
    return (entry.interno ?? "").trim() !== ""
        || (entry.user.mobilePhone ?? "").trim() !== ""
        || (entry.user.faxNumber ?? "").trim() !== "";
}

export interface IOpzioniVista {
    soloMembri: boolean;            // nasconde gli ospiti (userType Guest)
    nascondiSenzaMail: boolean;     // nasconde gli account senza mail
    nascondiSenzaTelefono: boolean; // nasconde gli utenti senza alcun numero
    includiEsterni: boolean;        // checkbox "Esterni"
}

// Applica i filtri di visualizzazione DOPO il merge: così un utente nascosto
// (es. un ospite) non fa ricomparire la sua voce di lista come "servizio".
export function applicaVista(rubrica: IRubricaEntry[], opz: IOpzioniVista): IRubricaEntry[] {
    return rubrica.filter(entry => {
        if (entry.tipo === "esterno") return opz.includiEsterni;

        if (entry.tipo === "utente") {
            if (opz.soloMembri && entry.user.userType === "Guest") return false;
            if (opz.nascondiSenzaMail && (entry.user.mail ?? "").trim() === "") return false;
            if (opz.nascondiSenzaTelefono && !haTelefono(entry)) return false;
        }

        return true;
    });
}

// minuscolo e senza accenti: "Qualità" e "qualita" diventano uguali
function normalizza(s?: string | null): string {
    return (s ?? "")
        .normalize("NFD")
        .replace(/[̀-ͯ]/g, "")
        .toLowerCase();
}

function soloCifre(s?: string | null): string {
    return (s ?? "").replace(/\D/g, "");
}

// parole che selezionano un intero tipo di voce (solo come parola intera)
const PAROLE_TIPO: { [tipo: string]: string[] } = {
    servizio: ["servizio", "servizi"],
    esterno: ["esterno", "esterni"],
};

// Campi in cui si cerca come testo (nomi, reparti, numeri BREVI): basta che la parola
// compaia in un punto qualsiasi.
function campiTesto(entry: IRubricaEntry): Array<string | null | undefined> {
    switch (entry.tipo) {
        case "utente":
            return [entry.user.displayName, entry.user.mail, entry.user.department, entry.user.jobTitle, entry.interno, entry.user.faxNumber];
        case "servizio":
            return [entry.nome, entry.reparto, entry.interno];
        default:
            return [entry.nome, entry.numeroBreve];
    }
}

// Numeri scritti per intero: si cercano solo per cifre (vedi MIN_CIFRE_NUMERO_LUNGO).
function numeriLunghi(entry: IRubricaEntry): string[] {
    const numeri = entry.tipo === "utente" ? [entry.user.mobilePhone]
        : entry.tipo === "esterno" ? [entry.numero]
        : [];
    return numeri.map(soloCifre).filter(c => c !== "");
}

// Tutti i numeri brevi della voce, uno per elemento (interni, numero breve del cellulare,
// numero breve degli esterni). Servono a capire se un numero è già assegnato.
function numeriBrevi(entry: IRubricaEntry): string[] {
    switch (entry.tipo) {
        case "utente": {
            const breve = (entry.user.faxNumber ?? "").trim();
            return dividiNumeri(entry.interno).concat(breve !== "" ? [breve] : []);
        }
        case "servizio":
            return dividiNumeri(entry.interno);
        default:
            return dividiNumeri(entry.numeroBreve);
    }
}

export function filtra(input: string, risultati: IRubricaEntry[]): IRubricaEntry[] {
    // più parole: ogni parola deve comparire in almeno un campo, in qualsiasi ordine
    const parole = normalizza(input).split(/\s+/).filter(p => p !== "");
    if (parole.length === 0) return []; // nulla mostrato finché non si cerca

    // Un numero digitato con spazi o prefisso ("+39 333 123 4567", "333 1234567") è fatto di
    // più "parole": lo si confronta anche per intero, senza separatori, con i numeri lunghi.
    const compatta = input.replace(/[\s.\-()+]/g, "");
    const numeroIntero = /^\d+$/.test(compatta) && compatta.length >= MIN_CIFRE_NUMERO_LUNGO ? compatta : "";

    return risultati.filter(entry => {
        const haystack = campiTesto(entry).map(t => normalizza(t)).join(" ");
        const lunghi = numeriLunghi(entry);

        const tutteLeParole = parole.every(p =>
            haystack.indexOf(p) !== -1 ||
            // "servizio"/"servizi" e "esterno"/"esterni" trovano tutte le voci di quel tipo
            (entry.tipo !== "utente" && PAROLE_TIPO[entry.tipo].indexOf(p) !== -1) ||
            // numero lungo (cellulare, numero esterno): solo con abbastanza cifre
            (p.length >= MIN_CIFRE_NUMERO_LUNGO && /^\d+$/.test(p) && lunghi.some(c => c.indexOf(p) !== -1))
        );

        return tutteLeParole || (numeroIntero !== "" && lunghi.some(c => c.indexOf(numeroIntero) !== -1));
    });
}

// --- numeri brevi liberi ---

export interface IIntervalloNumeri {
    min: number;
    max: number;
}

export interface INumeroBreveLibero extends IIntervalloNumeri {
    numero: string; // come l'ha digitato l'utente
}

// Intervallo dei numeri brevi degli esterni (minimo e massimo): è DINAMICO, cambia
// con i dati della lista. Null se nessun esterno ha un numero breve numerico.
export function intervalloNumeriBreviEsterni(rubrica: IRubricaEntry[]): IIntervalloNumeri | null {
    let min = Infinity;
    let max = -Infinity;

    for (const entry of rubrica) {
        if (entry.tipo !== "esterno") continue;
        for (const n of dividiNumeri(entry.numeroBreve)) {
            if (!/^\d+$/.test(n)) continue;
            const v = Number(n);
            if (v < min) min = v;
            if (v > max) max = v;
        }
    }

    return min === Infinity ? null : { min, max };
}

// Se "input" è un numero che rientra nell'intervallo dei numeri brevi degli esterni ma
// non è assegnato a NESSUNA voce (né esterno, né utente, né servizio) restituisce
// l'intervallo, altrimenti null. "rubrica" deve essere l'elenco completo, senza filtri:
// un numero di un utente nascosto (es. ospite) è comunque assegnato.
export function numeroBreveNonAssegnato(input: string, rubrica: IRubricaEntry[]): INumeroBreveLibero | null {
    const numero = input.trim();
    if (!/^\d+$/.test(numero)) return null;

    const intervallo = intervalloNumeriBreviEsterni(rubrica);
    if (!intervallo) return null;

    const valore = Number(numero);
    if (valore < intervallo.min || valore > intervallo.max) return null;

    const assegnato = rubrica.some(entry =>
        numeriBrevi(entry).some(n => /^\d+$/.test(n) && Number(n) === valore)
    );

    return assegnato ? null : { numero, min: intervallo.min, max: intervallo.max };
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
