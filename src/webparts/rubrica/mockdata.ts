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
    PREFISSO_INTERNO: "00",

    // Range interni: da 2300 a 2599  (300 numeri disponibili)
    INTERNO_BASE: 0,
    INTERNO_RANGE: 300,

    // Quanti utenti Graph generare
    NUM_UTENTI: 150,

    // Quanti servizi/uffici aggiuntivi (senza account AD)
    NUM_SERVIZI: 80,

    // Percentuale di utenti che hanno un numero interno (0–1)
    COPERTURA_INTERNI_UTENTI: 0.80,

    // Seed deterministico (cambia per avere combinazioni diverse)
    SEED: 42,
};

// ─────────────────────────────────────────────────────────────────────────────
//  📦  IMPORT — adatta i percorsi al tuo progetto se necessario
// ─────────────────────────────────────────────────────────────────────────────

import { User } from "@microsoft/microsoft-graph-types";
import { IContattoInterno } from "./types"; // ← percorso relativo standard

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
    "Alberico", "Brunella", "Celestino", "Dorotea", "Ermenegildo",
    "Fioralba", "Goffredo", "Hildegarda", "Imerio", "Jole",
    "Katiuscia", "Learco", "Medea", "Nereo", "Orsola",
    "Pieralberto", "Quirino", "Rosmunda", "Saverio", "Teodolinda",
    "Ubaldo", "Venanzio", "Walburga", "Xanto", "Ylenia",
    "Zaccaria", "Adalgisa", "Beniamino", "Clotilde", "Dagoberto",
    "Eleuteria", "Fulgenzio", "Gelsomina", "Hannibal", "Isidora",
    "Ladislao", "Mansueto", "Natanaele", "Onofrio", "Perpetua",
    "Reginaldo", "Sigismondo", "Tarcisio", "Ugolino", "Virgilio",
    "Wanda", "Zenobio", "Arcibaldo", "Baldassarre", "Crispino",
];

const COGNOMI = [
    "Brassoletti", "Cazzutti", "Dinaccioli", "Florimondi", "Garbuglia",
    "Herculani", "Intragnoli", "Lapuzza", "Morbiducci", "Norcenni",
    "Orzatelli", "Pampanini", "Quaresimotti", "Ruffatazzo", "Sbrocchetti",
    "Trivellacci", "Ubertazzi", "Volpacchiotti", "Wronkowski", "Zabaioni",
    "Andrisani", "Bertaccini", "Capannacci", "Delbuono", "Ercolessi",
    "Farfallini", "Guarnieri", "Innamorati", "Lionetti", "Maccagnani",
    "Nasturzetti", "Odorisio", "Pantaleoni", "Raccuja", "Scalambrone",
    "Tagliacozzo", "Ugoccioni", "Verdicchietti", "Zampetti", "Acciarino",
    "Balducci", "Ciampiconi", "Durantini", "Espurini", "Fiordimonte",
    "Giroldetti", "Iannaccone", "Lambertucci", "Mastrobuoni", "Negrisoli",
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
    "Centralino", "Portineria", "Sala Riunioni A", "Sala Riunioni B", "Sala Riunioni C",
    "Fax Amministrazione", "Fax Legale", "Stampante Piano 1", "Stampante Piano 2",
    "Archivio Generale", "Archivio Storico", "Magazzino Centrale", "Magazzino Secondario",
    "Infermeria", "Reception Nord", "Reception Sud", "Cancello Ingresso",
    "Ufficio Acquisti", "Ufficio Spedizioni", "Ufficio Reclami",
    "Helpdesk IT", "Server Room", "NOC",
    "Bar Aziendale", "Mensa Aziendale", "Sala Relax",
    "Autorimessa", "Guardiola Notte", "Servizio Pulizie",
    "Ufficio Stampa", "Sala Conferenze Executive", "Sala Videoconferenze",
    "Centro Formazione", "Biblioteca Tecnica", "Laboratorio Test",
    "Ufficio Viaggi", "Gestione Fleet", "Ufficio Sinistri",
    "Poliambulatorio Aziendale", "Ufficio Sicurezza", "RSPP",
    "Centralino Emergenze", "Portiere Notturno", "Reparto Spedizioni Urgenti",
    "Assistenza Clienti", "Servizio Post-Vendita", "Ufficio Resi",
    "Ufficio Brevetti", "Ufficio Gare e Appalti", "Segreteria Direzione",
    "Segreteria Commerciale", "Segreteria Tecnica", "Corrispondenza Interna",
    "Reparto Produzione A", "Reparto Produzione B", "Controllo Qualità Linea",
    "Manutenzione Impianti", "Officina Interna", "Magazzino Ricambi",
    "Centralino Assistenza Tecnica", "Pronto Intervento Tecnico",
    "Accettazione Merci", "Controllo Accessi", "Ufficio Export",
    "Ufficio Import", "Ufficio Doganale", "Ufficio Assicurazioni",
    "Back Office Commerciale", "Gestione Ordini", "Customer Success",
    "Ufficio Sostenibilità", "Ufficio ESG", "Ufficio Privacy",
    "DPO – Data Protection Officer", "Compliance", "Internal Audit",
    "Investor Relations", "Ufficio Finanziamenti", "Tesoreria",
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
function generaUtenti(): User[] {
    // Assicura coppie nome+cognome uniche
    const usate = new Set<string>();
    const utenti: User[] = [];

    let tentativi = 0;
    while (utenti.length < CFG.NUM_UTENTI && tentativi < 10000) {
        tentativi++;
        const nome = pick(NOMI);
        const cognome = pick(COGNOMI);
        const key = `${nome}_${cognome}`;
        if (usate.has(key)) continue;
        usate.add(key);

        const reparto = pick(REPARTI);
        const mail = `${slugify(nome)}.${slugify(cognome)}@${CFG.DOMAIN}`;
        const id = `mock-uid-${utenti.length + 1}-${nome.slice(0, 3).toLowerCase()}`;

        utenti.push({
            id,
            displayName: `${nome} ${cognome}`,
            mail,
            department: reparto.dept,
            jobTitle: pick(reparto.jobs),
            mobilePhone: `+39 3${Math.floor(rand() * 9)}${("0000000" + Math.floor(rand() * 10000000)).slice(-7)}`,
            businessPhones: [],
            accountEnabled: true,
        } as User);
    }

    return utenti;
}

/** Genera le voci lista SharePoint */
function generaVociLista(utenti: User[]): IContattoInterno[] {
    const voci: IContattoInterno[] = [];
    const interniUsati = new Set<string>();

    function nuovoInterno(): string {
        let i: string;
        let loop = 0;
        do { i = interno(); loop++; } while (interniUsati.has(i) && loop < 1000);
        interniUsati.add(i);
        return i;
    }

    // Quanti utenti hanno un interno
    const numConInterno = Math.round(utenti.length * CFG.COPERTURA_INTERNI_UTENTI);
    const utentiConInterno = pickN(utenti, numConInterno);

    for (const u of utentiConInterno) {
        voci.push({ contatto: u.mail!, interno: nuovoInterno() });
    }

    // Servizi aggiuntivi
    const serviziScelti = pickN(NOMI_SERVIZI, CFG.NUM_SERVIZI);
    for (const nome of serviziScelti) {
        voci.push({ contatto: nome, interno: nuovoInterno() });
    }

    return voci;
}

// ─────────────────────────────────────────────────────────────────────────────
//  📤  EXPORT — usa questi al posto di getGraphUsers / getVociLista
// ─────────────────────────────────────────────────────────────────────────────

// Generazione una-tantum (singleton, dati stabili per tutta la sessione)
const _utenti: User[] = generaUtenti();
const _voci: IContattoInterno[] = generaVociLista(_utenti);

/** Simula getGraphUsers() — aggiunge 300ms di latenza finta */
export async function getMockGraphUsers(): Promise<User[]> {
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

// console.table(_utenti.slice(0, 10).map(u => ({ name: u.displayName, mail: u.mail, dept: u.department })));
// console.table(_voci.slice(0, 10));
// console.log(`Utenti: ${_utenti.length} | Voci lista: ${_voci.length}`);