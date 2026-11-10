// =============================================================================
//  ██████╗ ██╗   ██╗██████╗ ██████╗ ██╗ ██████╗ █████╗
//  ██╔══██╗██║   ██║██╔══██╗██╔══██╗██║██╔════╝██╔══██╗
//  ██████╔╝██║   ██║██████╔╝██████╔╝██║██║     ███████║
//  ██╔══██╗██║   ██║██╔══██╗██╔══██╗██║██║     ██╔══██║
//  ██║  ██║╚██████╔╝██████╔╝██║  ██║██║╚██████╗██║  ██║
//  ╚═╝  ╚═╝ ╚═════╝ ╚═════╝ ╚═╝  ╚═╝╚═╝ ╚═════╝╚═╝  ╚═╝
//  Generatore dati finti per sviluppo e test locale
//  Sostituisce getGraphUsers() e getVociLista() con dati deterministici.
// =============================================================================

// ─────────────────────────────────────────────────────────────────────────────
//  🔧  CONFIGURAZIONE  — modifica qui, non toccare il resto
// ─────────────────────────────────────────────────────────────────────────────
const CFG = {
    // Dominio email aziendale  →  mario.rossi@QUESTO.it
    DOMAIN: "coind.it",

    // Prefisso interni telefonici  →  "02" → interni tipo "02-2301"
    PREFISSO_INTERNO: "",

    // Range interni: da 2300 a 2599  (300 numeri disponibili)
    INTERNO_BASE: 100,
    INTERNO_RANGE: 300,

    // Quanti utenti Graph generare
    NUM_UTENTI: 150,

    // Quanti servizi/uffici aggiuntivi (senza account AD)
    NUM_SERVIZI: 80,

    // Quanti contatti esterni generare (lista "Esterni")
    NUM_ESTERNI: 25,

    // Percentuale di utenti che hanno un numero interno (0–1)
    COPERTURA_INTERNI_UTENTI: 0.80,

    // Seed deterministico (cambia per avere combinazioni diverse)
    SEED: 42,
};

// ─────────────────────────────────────────────────────────────────────────────
//  📦  IMPORT — adatta i percorsi al tuo progetto se necessario
// ─────────────────────────────────────────────────────────────────────────────

import { IContattoEsterno, IContattoInterno, UtenteGraph } from "./types"; // ← percorso relativo standard

// ─────────────────────────────────────────────────────────────────────────────
//  🎲  PRNG deterministico (mulberry32) — stesso seed = stessi dati sempre
// ─────────────────────────────────────────────────────────────────────────────

function makePrng(seed: number) {
    let s = seed;
    return () => {
        s |= 0;
        s = (s + 0x6d2b79f5) | 0;
        let t = Math.imul(s ^ (s >>> 15), 1 | s);
        t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

const rand = makePrng(CFG.SEED);

function pick<T>(arr: T[]): T {
    return arr[Math.floor(rand() * arr.length)];
}

function pickN<T>(arr: T[], n: number): T[] {
    const shuffled = [...arr].sort(() => rand() - 0.5);
    return shuffled.slice(0, n);
}

function interno(): string {
    return `${CFG.PREFISSO_INTERNO}${CFG.INTERNO_BASE + Math.floor(rand() * CFG.INTERNO_RANGE)}`;
}

// ─────────────────────────────────────────────────────────────────────────────
//  👤  POOL NOMI
//  Palesemente inventati (suffissi -etti, -azzo, -ucci…) ma foneticamente italiani
// ─────────────────────────────────────────────────────────────────────────────

const NOMI = [
    "Marco", "Laura", "Giuseppe", "Francesca", "Antonio",
    "Giulia", "Matteo", "Sara", "Andrea", "Chiara",
    "Luca", "Valentina", "Roberto", "Alessia", "Stefano",
    "Martina", "Davide", "Elena", "Simone", "Alice",
    "Fabio", "Federica", "Giovanni", "Paola", "Riccardo",
    "Silvia", "Nicola", "Monica", "Alessandro", "Serena",
    "Paolo", "Cristina", "Enrico", "Angela", "Michele",
    "Elisa", "Filippo", "Veronica", "Lorenzo", "Roberta",
    "Giorgio", "Claudia", "Daniele", "Ilaria", "Francesco",
    "Michela", "Emanuele", "Sabrina", "Massimo", "Giorgia",
];

const COGNOMI = [
    "Rossi", "Ferrari", "Russo", "Bianchi", "Romano",
    "Colombo", "Ricci", "Marino", "Greco", "Bruno",
    "Gallo", "Conti", "De Luca", "Mancini", "Costa",
    "Giordano", "Rizzo", "Lombardi", "Moretti", "Barbieri",
    "Fontana", "Santoro", "Marini", "Rinaldi", "Caruso",
    "Ferretti", "Gatti", "Battaglia", "Pellegrini", "Palumbo",
    "Sartori", "Fabbri", "Villa", "Coppola", "Martinelli",
    "Serra", "Cattaneo", "Longo", "Gentile", "De Angelis",
    "Benedetti", "Barone", "Montanari", "Ferraro", "Vitale",
    "Neri", "Martini", "Sala", "Bernardi", "Testa",
];
// ─────────────────────────────────────────────────────────────────────────────
//  🏢  REPARTI & RUOLI
// ─────────────────────────────────────────────────────────────────────────────

const REPARTI: { dept: string; jobs: string[] }[] = [
    {
        dept: "Amministrazione",
        jobs: ["Responsabile Amministrativo", "Addetto Contabilità", "Controller", "Impiegato Amministrativo"],
    },
    {
        dept: "Risorse Umane",
        jobs: ["HR Manager", "Addetto Selezione", "HR Business Partner", "Responsabile Paghe"],
    },
    {
        dept: "IT",
        jobs: ["Sviluppatore Senior", "Sistemista", "DevOps Engineer", "IT Support", "Analista Funzionale"],
    },
    {
        dept: "Commerciale",
        jobs: ["Area Manager", "Account Executive", "Inside Sales", "Key Account Manager"],
    },
    {
        dept: "Marketing",
        jobs: ["Marketing Manager", "Social Media Specialist", "Copywriter", "Graphic Designer"],
    },
    {
        dept: "Logistica",
        jobs: ["Responsabile Magazzino", "Operatore Logistico", "Spedizioniere", "Coordinatore Supply Chain"],
    },
    {
        dept: "Legale",
        jobs: ["Responsabile Legale", "Consulente Legale", "Paralegal"],
    },
    {
        dept: "Qualità",
        jobs: ["Responsabile Qualità", "Auditor Interno", "Tecnico Qualità"],
    },
    {
        dept: "Direzione",
        jobs: ["Direttore Generale", "CFO", "CTO", "COO", "Assistente di Direzione"],
    },
    {
        dept: "R&D",
        jobs: ["Ricercatore", "Ingegnere R&D", "Tecnico di Laboratorio", "Project Manager R&D"],
    },
];

// ─────────────────────────────────────────────────────────────────────────────
//  📞  SERVIZI / UFFICI senza account AD
// ─────────────────────────────────────────────────────────────────────────────

const NOMI_SERVIZI = [
    "Centralino", "Portineria", "Reception",
    "Magazzino Materie Prime", "Magazzino Prodotti Finiti", "Magazzino Imballaggi",
    "Accettazione Merci", "Spedizioni", "Reparto Spedizioni Urgenti",
    "Reparto Tostatura", "Reparto Macinazione", "Reparto Confezionamento",
    "Reparto Porzionato", "Reparto Miscele", "Controllo Qualità Linea",
    "Laboratorio Analisi Sensoriale", "Laboratorio Chimico", "Camera di Degustazione",
    "Manutenzione Impianti", "Officina Interna", "Magazzino Ricambi",
    "Sala Caldaie", "Sala Compressori", "Impianto Trattamento Aria",
    "Ufficio Acquisti Materie Prime", "Ufficio Import Verde", "Ufficio Doganale",
    "Ufficio Export", "Gestione Ordini", "Back Office Commerciale",
    "Assistenza Clienti", "Servizio Post-Vendita", "Ufficio Reclami",
    "Helpdesk IT", "Server Room",
    "Ufficio Sicurezza", "Infermeria",
    "Mensa Aziendale", "Bar Aziendale",
    "Archivio Generale", "Segreteria Direzione",
    "Fax Amministrazione", "Stampante Piano 1", "Stampante Piano 2",
    "Guardiola Notte", "Portiere Notturno", "Autorimessa",
    "Centralino Emergenze", "Pronto Intervento Tecnico",
];

// ─────────────────────────────────────────────────────────────────────────────
//  🔧  GENERATORI
// ─────────────────────────────────────────────────────────────────────────────

function slugify(s: string): string {
    return s
        .toLowerCase()
        .replace(/à/g, "a").replace(/è|é/g, "e").replace(/ì/g, "i")
        .replace(/ò/g, "o").replace(/ù/g, "u")
        .replace(/[^a-z0-9]/g, ".");
}

/** Genera un array di User Graph finti */
function generaUtenti(): UtenteGraph[] {
    // Assicura coppie nome+cognome uniche
    const usate = new Set<string>();
    const utenti: UtenteGraph[] = [];

    let tentativi = 0;
    while (utenti.length < CFG.NUM_UTENTI && tentativi < 10000) {
        tentativi++;
        const nome = pick(NOMI);
        const cognome = pick(COGNOMI);
        const key = `${nome}_${cognome}`;
        if (usate.has(key)) continue;
        usate.add(key);

        const reparto = pick(REPARTI);
        const ruolo = pick(reparto.jobs); // stesso ordine di estrazione di prima
        const mail = `${slugify(nome)}.${slugify(cognome)}@${CFG.DOMAIN}`;
        const id = `mock-uid-${utenti.length + 1}-${nome.slice(0, 3).toLowerCase()}`;

        // il PRNG si consuma sempre, anche quando il cellulare poi non viene assegnato
        const cellulare = `+39 3${Math.floor(rand() * 9)}${("0000000" + Math.floor(rand() * 10000000)).slice(-7)}`;
        const n = utenti.length + 1;

        utenti.push({
            id,
            displayName: `${nome} ${cognome}`,
            mail,
            // come in AD: department e jobTitle, mostrati come "Reparto - Job title"
            department: reparto.dept,
            jobTitle: ruolo,
            // cellulare aziendale: un utente ogni 5 non lo ha
            mobilePhone: n % 5 === 0 ? undefined : cellulare,
            // numero breve del cellulare (in AD: faxNumber): solo per metà dei cellulari,
            // nel range 400-551 (unici per n < 152, e fuori da quello degli interni)
            faxNumber: n % 5 !== 0 && n % 2 === 0 ? String(400 + ((n * 53) % 152)) : undefined,
            // numero breve = ultime 3 cifre; un utente ogni 9 non ha il telefono
            businessPhones: n % 9 === 0
                ? []
                : [`+39 051 6543${("00" + ((n * 37) % 1000)).slice(-3)}`],
            accountEnabled: true,
            // ogni 13esimo utente è un ospite (nessun uso del PRNG: gli altri dati restano identici)
            userType: n % 13 === 0 ? "Guest" : "Member",
        } as UtenteGraph);
    }

    return utenti;
}

// Interni già assegnati (condivisi tra servizi ed esterni, così non si ripetono)
const interniUsati = new Set<string>();

function nuovoInterno(): string {
    let i: string;
    let loop = 0;
    do { i = interno(); loop++; } while (interniUsati.has(i) && loop < 1000);
    interniUsati.add(i);
    return i;
}

// Colonna "Reparto" della lista servizi, ricavata dal nome (nessun uso del PRNG).
// Per alcuni servizi resta vuoto: così si prova anche il caso "reparto assente".
function repartoServizio(nome: string): string {
    if (/magazzino|spedizion|accettazione|autorimessa/i.test(nome)) return "Logistica";
    if (/reparto|manutenzione|officina|sala|impianto|ricambi|pronto intervento/i.test(nome)) return "Produzione";
    if (/laborator|qualit|degustazione/i.test(nome)) return "Qualità";
    if (/acquisti|import|doganale|export|ordini|commerciale|clienti|vendita|reclami/i.test(nome)) return "Commerciale";
    if (/helpdesk|server|stampante|fax/i.test(nome)) return "IT";
    return ""; // centralino, mensa, portineria...
}

/** Genera le voci lista SharePoint */
function generaVociLista(utenti: UtenteGraph[]): IContattoInterno[] {
    const voci: IContattoInterno[] = [];

    // Quanti utenti hanno un interno
    const numConInterno = Math.round(utenti.length * CFG.COPERTURA_INTERNI_UTENTI);
    const utentiConInterno = pickN(utenti, numConInterno);

    for (const u of utentiConInterno) {
        voci.push({ contatto: u.mail!, interno: nuovoInterno() });
    }

    // Servizi aggiuntivi
    const serviziScelti = pickN(NOMI_SERVIZI, CFG.NUM_SERVIZI);
    for (const nome of serviziScelti) {
        voci.push({ contatto: nome, interno: nuovoInterno(), reparto: repartoServizio(nome) });
    }

    return voci;
}

// ─────────────────────────────────────────────────────────────────────────────
//  📤  EXPORT — usa questi al posto di getGraphUsers / getVociLista
// ─────────────────────────────────────────────────────────────────────────────

// Generazione una-tantum (singleton, dati stabili per tutta la sessione)
const _utenti: UtenteGraph[] = generaUtenti();
const _voci: IContattoInterno[] = generaVociLista(_utenti);

/** Simula getGraphUsers() — aggiunge 300ms di latenza finta */
export async function getMockGraphUsers(): Promise<UtenteGraph[]> {
    await new Promise((resolve) => setTimeout(resolve, 300));
    return _utenti;
}

/** Simula getVociLista() — aggiunge 200ms di latenza finta */
export async function getMockVociLista(): Promise<IContattoInterno[]> {
    await new Promise((resolve) => setTimeout(resolve, 200));
    return _voci;
}

// ─────────────────────────────────────────────────────────────────────────────
//  🔍  DEBUG — decommentare in console per ispezionare i dati
// ─────────────────────────────────────────────────────────────────────────────

// console.table(_utenti.slice(0, 10).map(u => ({ name: u.displayName, mail: u.mail, job: u.jobTitle })));
// console.table(_voci.slice(0, 10));
// console.log(`Utenti: ${_utenti.length} | Voci lista: ${_voci.length}`);

// ─────────────────────────────────────────────────────────────────────────────
//  🤝  ESTERNI (lista "Esterni") + account particolari per provare i filtri
//  Generati DOPO tutto il resto: i dati sopra non cambiano rispetto a prima.
// ─────────────────────────────────────────────────────────────────────────────

const NOMI_ESTERNI = [
    "Studio Commercialista Bianchi", "Consulente Sicurezza Verdi", "Idraulico Pronto Intervento",
    "Elettricista Neri", "Tecnico Caldaie Fontana", "Assistenza Fotocopiatrici",
    "Corriere Espresso", "Taxi Navile", "Studio Legale Marini",
    "Medico Competente", "Società Pulizie", "Manutenzione Ascensori",
    "Agenzia Interinale", "Consulente Paghe Esterno", "Fornitore Caffè Verde",
    "Fornitore Imballaggi", "Trasportatore Nord", "Trasportatore Sud",
    "Revisore dei Conti", "Consulente Privacy", "Fornitore Etichette",
    "Assistenza Software Gestionale", "Assistenza Linea Confezionamento",
    "Ditta Disinfestazione", "Centro Analisi Esterno",
];

/** Genera le voci della lista "Esterni": Nome + Numero (intero) + Numero breve */
function generaEsterni(): IContattoEsterno[] {
    return pickN(NOMI_ESTERNI, CFG.NUM_ESTERNI).map((nome, i) => {
        const numero = `+39 051 ${Math.floor(rand() * 9000000) + 1000000}`;
        return {
            nome,
            // uno ogni 7 non ha il numero breve, uno ogni 11 non ha il numero intero
            numero: (i + 1) % 11 === 0 ? "" : numero,
            numeroBreve: (i + 1) % 7 === 0 ? "" : nuovoInterno(),
        };
    });
}

const _esterni: IContattoEsterno[] = generaEsterni();

// Account senza mail (es. sale riunioni): servono a provare "Nascondi gli account senza mail".
// Aggiunti dopo la generazione della lista, quindi non hanno interno.
_utenti.push(
    { id: "mock-uid-sala-1", displayName: "Sala Riunioni 1", jobTitle: "Sala", userType: "Member", accountEnabled: true } as UtenteGraph,
    { id: "mock-uid-sala-2", displayName: "Sala Riunioni 2", jobTitle: "Sala", userType: "Member", accountEnabled: true } as UtenteGraph,
);

// Account con mail ma senza alcun numero: servono a provare "Nascondi gli utenti senza alcun numero".
// Il terzo ha solo il numero breve del cellulare, quindi resta sempre visibile.
_utenti.push(
    { id: "mock-uid-notel-1", displayName: "Zeno Senzatelefono", mail: `zeno.senzatelefono@${CFG.DOMAIN}`, jobTitle: "Consulente", userType: "Member", accountEnabled: true } as UtenteGraph,
    { id: "mock-uid-notel-2", displayName: "Zara Senzanumeri", mail: `zara.senzanumeri@${CFG.DOMAIN}`, jobTitle: "Stagista", businessPhones: [], userType: "Member", accountEnabled: true } as UtenteGraph,
    { id: "mock-uid-breve-1", displayName: "Zoe Solobreve", mail: `zoe.solobreve@${CFG.DOMAIN}`, jobTitle: "Tecnico", faxNumber: "555", userType: "Member", accountEnabled: true } as UtenteGraph,
);

/** Simula getVociLista() sulla lista Esterni — aggiunge 250ms di latenza finta */
export async function getMockEsterni(): Promise<IContattoEsterno[]> {
    await new Promise((resolve) => setTimeout(resolve, 250));
    return _esterni;
}
