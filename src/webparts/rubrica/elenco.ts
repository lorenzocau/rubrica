// Dati per la vista Elenco (tabelle da stampare e da esportare in CSV).
// Funzioni pure, senza dipendenze da SPFx né dal DOM: si possono testare.
//
// Tre tabelle con colonne diverse:
//   Dipendenti: Cognome Nome | N. interno breve | Cellulare | Cellulare breve | Job title
//   Servizi:    Nome | Reparto | Interno
//   Esterni:    Nome | Numero breve | Numero
//
// Regole fisse di questa vista (diverse da quelle della Ricerca):
//   - ospiti (Guest) e account senza mail non compaiono mai;
//   - gli utenti senza alcun numero compaiono comunque, con le celle vuote, a meno che
//     l'opzione "Nascondi gli utenti senza numero interno breve" del pannello non li tolga.

import { IRubricaEntry, IRubricaEsterno, IRubricaServizio, IRubricaUtente, UtenteGraph } from "./types";
import { applicaVista, filtra } from "./rubricaLogic";

export type TipoTabella = "dipendenti" | "servizi" | "esterni";
export type ChiaveColonna = "nome" | "interno" | "cellulare" | "cellulareBreve" | "jobTitle" | "reparto";

export interface IColonnaTabella {
    chiave: ChiaveColonna;
    titolo: string;
}

export type RigaTabella = { [K in ChiaveColonna]?: string };

export interface ITabella {
    id: TipoTabella;
    titolo: string;   // titolo della sezione ("Dipendenti")
    tipo: string;     // valore della colonna "Tipo" nel CSV ("Dipendente")
    colonne: IColonnaTabella[];
    righe: RigaTabella[]; // contengono solo le chiavi delle colonne visibili
}

export interface IOpzioniTabelle {
    dipendenti: boolean;     // sezione attiva
    servizi: boolean;
    esterni: boolean;
    nascondiSenzaInterno: boolean; // come nel pannello: nasconde i dipendenti senza numero interno breve
    mostraCellulare: boolean; // come nel pannello: nasconde cellulare e cellulare breve
    mostraReparto: boolean;   // come nel pannello: nasconde job title e reparto
}

const collator = new Intl.Collator("it", { sensitivity: "base" });

function pulito(s?: string | null): string {
    return (s ?? "").trim();
}

// "Cognome Nome" da surname e givenName di Graph.
// Se manca il cognome non si può ordinare per cognome: si ripiega sul nome visualizzato
// (o sulla mail), così anche gli account incompleti compaiono con un nome leggibile.
export function nomeDipendente(user: UtenteGraph): string {
    const cognome = pulito(user.surname);
    const nome = pulito(user.givenName);

    if (cognome === "") {
        const visualizzato = pulito(user.displayName);
        return visualizzato !== "" ? visualizzato : pulito(user.mail);
    }
    return nome === "" ? cognome : `${cognome} ${nome}`;
}

// Ordine alfabetico per la chiave "nome", senza distinguere maiuscole e accenti;
// a parità di nome resta l'ordine di arrivo.
function ordinaPerNome(righe: RigaTabella[]): RigaTabella[] {
    return righe
        .map((riga, indice) => ({ riga, indice }))
        .sort((a, b) => collator.compare(a.riga.nome ?? "", b.riga.nome ?? "") || a.indice - b.indice)
        .map(x => x.riga);
}

function tabellaDipendenti(utenti: IRubricaUtente[], opz: IOpzioniTabelle): ITabella {
    const colonne: IColonnaTabella[] = [
        { chiave: "nome", titolo: "Cognome Nome" },
        { chiave: "interno", titolo: "N. interno breve" },
    ];
    if (opz.mostraCellulare) {
        colonne.push({ chiave: "cellulare", titolo: "Cellulare" });
        colonne.push({ chiave: "cellulareBreve", titolo: "Cellulare breve" });
    }
    if (opz.mostraReparto) colonne.push({ chiave: "jobTitle", titolo: "Job title" });

    const righe = utenti.map(u => {
        const riga: RigaTabella = { nome: nomeDipendente(u.user), interno: pulito(u.interno) };
        if (opz.mostraCellulare) {
            riga.cellulare = pulito(u.user.mobilePhone);
            riga.cellulareBreve = pulito(u.user.faxNumber);
        }
        if (opz.mostraReparto) riga.jobTitle = pulito(u.user.jobTitle);
        return riga;
    });

    return { id: "dipendenti", titolo: "Dipendenti", tipo: "Dipendente", colonne, righe: ordinaPerNome(righe) };
}

function tabellaServizi(servizi: IRubricaServizio[], opz: IOpzioniTabelle): ITabella {
    const colonne: IColonnaTabella[] = [{ chiave: "nome", titolo: "Nome" }];
    if (opz.mostraReparto) colonne.push({ chiave: "reparto", titolo: "Reparto" });
    colonne.push({ chiave: "interno", titolo: "Interno" });

    const righe = servizi.map(s => {
        const riga: RigaTabella = { nome: s.nome, interno: s.interno };
        if (opz.mostraReparto) riga.reparto = s.reparto;
        return riga;
    });

    return { id: "servizi", titolo: "Servizi", tipo: "Servizio", colonne, righe: ordinaPerNome(righe) };
}

function tabellaEsterni(esterni: IRubricaEsterno[]): ITabella {
    const colonne: IColonnaTabella[] = [
        { chiave: "nome", titolo: "Nome" },
        { chiave: "interno", titolo: "Numero breve" },
        { chiave: "cellulare", titolo: "Numero" },
    ];
    const righe = esterni.map(e => ({ nome: e.nome, interno: e.numeroBreve, cellulare: e.numero }));

    return { id: "esterni", titolo: "Esterni", tipo: "Esterno", colonne, righe: ordinaPerNome(righe) };
}

// "rubrica" è l'elenco completo (senza filtri). "ricerca" vuota = tutto.
// Restituisce solo le sezioni attive in "opz", anche se senza righe.
export function costruisciTabelle(rubrica: IRubricaEntry[], ricerca: string, opz: IOpzioniTabelle): ITabella[] {
    const visibili = applicaVista(rubrica, {
        soloMembri: true,
        nascondiSenzaMail: true,
        nascondiSenzaTelefono: false,
        nascondiSenzaInterno: opz.nascondiSenzaInterno,
        includiEsterni: true,
    });

    const cerca = ricerca.trim();
    const risultati = cerca === "" ? visibili : filtra(cerca, visibili);

    const tabelle: ITabella[] = [];
    if (opz.dipendenti) {
        tabelle.push(tabellaDipendenti(risultati.filter((e): e is IRubricaUtente => e.tipo === "utente"), opz));
    }
    if (opz.servizi) {
        tabelle.push(tabellaServizi(risultati.filter((e): e is IRubricaServizio => e.tipo === "servizio"), opz));
    }
    if (opz.esterni) {
        tabelle.push(tabellaEsterni(risultati.filter((e): e is IRubricaEsterno => e.tipo === "esterno")));
    }
    return tabelle;
}

// --- CSV ---
// Separatore ";" e righe CRLF: è il formato che Excel italiano apre con un doppio clic.
// Il BOM UTF-8 (per gli accenti) lo aggiunge chi scrive il file.

// Colonne del CSV nell'ordine fisso; compaiono solo quelle presenti in almeno una tabella.
// Per gli esterni "N. interno breve" è il numero breve e "Cellulare" il numero intero,
// come nella scheda.
const COLONNE_CSV: IColonnaTabella[] = [
    { chiave: "nome", titolo: "Nome" },
    { chiave: "interno", titolo: "N. interno breve" },
    { chiave: "cellulare", titolo: "Cellulare" },
    { chiave: "cellulareBreve", titolo: "Cellulare breve" },
    { chiave: "jobTitle", titolo: "Job title" },
    { chiave: "reparto", titolo: "Reparto" },
];

const CHIAVI_NUMERI: ChiaveColonna[] = ["interno", "cellulare", "cellulareBreve"];

// Un campo del CSV. Due protezioni per Excel:
//  - i numeri ("007", "+39 333 1234567") sono scritti come ="..." così non perdono gli zeri
//    iniziali e il "+" non viene preso per una formula (valgono solo valori fatti di cifre
//    e separatori, quindi non c'è nulla di eseguibile);
//  - qualsiasi altro testo che inizia con = + - @ riceve un apostrofo davanti, così non
//    diventa una formula.
export function campoCsv(valore: string, comeNumero: boolean): string {
    if (valore === "") return "";

    let v = valore;
    if (comeNumero && /^[+0-9][0-9 ()./-]*$/.test(v)) v = `="${v}"`;
    else if (/^[=+@\t\r-]/.test(v)) v = `'${v}`;

    return /[;"\r\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
}

export function generaCsv(tabelle: ITabella[]): string {
    const usate = COLONNE_CSV.filter(c => tabelle.some(t => t.colonne.some(x => x.chiave === c.chiave)));

    const righe: string[] = [
        ["Tipo"].concat(usate.map(c => c.titolo)).map(h => campoCsv(h, false)).join(";"),
    ];

    for (const tabella of tabelle) {
        for (const riga of tabella.righe) {
            const campi = [campoCsv(tabella.tipo, false)].concat(
                usate.map(c => campoCsv(riga[c.chiave] ?? "", CHIAVI_NUMERI.indexOf(c.chiave) !== -1))
            );
            righe.push(campi.join(";"));
        }
    }

    return righe.join("\r\n") + "\r\n";
}

// --- date ---

function due(n: number): string {
    return n < 10 ? `0${n}` : String(n);
}

// "16/11/2026 14:52"
export function formattaDataStampa(d: Date): string {
    return `${due(d.getDate())}/${due(d.getMonth() + 1)}/${d.getFullYear()} ${due(d.getHours())}:${due(d.getMinutes())}`;
}

// "rubrica_2026-11-16.csv"
export function nomeFileCsv(d: Date): string {
    return `rubrica_${d.getFullYear()}-${due(d.getMonth() + 1)}-${due(d.getDate())}.csv`;
}

// --- stampa ---

export function escapeHtml(s: string): string {
    return s
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

export interface IInfoStampa {
    titolo: string;     // titolo del documento ("Rubrica")
    dataStampa: string; // già formattata (vedi formattaDataStampa)
    filtro: string;     // testo cercato ("" = elenco completo): se c'è, lo si scrive in testa
}

function htmlTabella(tabella: ITabella): string {
    const intestazione = tabella.colonne.map(c => `<th>${escapeHtml(c.titolo)}</th>`).join("");
    const righe = tabella.righe
        .map(riga => {
            const celle = tabella.colonne
                .map(c => {
                    const classe = CHIAVI_NUMERI.indexOf(c.chiave) !== -1 ? ' class="num"' : "";
                    return `<td${classe}>${escapeHtml(riga[c.chiave] ?? "")}</td>`;
                })
                .join("");
            return `<tr>${celle}</tr>`;
        })
        .join("");

    return `<section><h2>${escapeHtml(tabella.titolo)} <span>(${tabella.righe.length})</span></h2>` +
        `<table><thead><tr>${intestazione}</tr></thead><tbody>${righe}</tbody></table></section>`;
}

const STILE_STAMPA = `
  @page { margin: 14mm 12mm; @bottom-center { content: "Pagina " counter(page) " di " counter(pages); font: 9pt Arial, sans-serif; color: #605e5c; } }
  * { box-sizing: border-box; }
  body { font: 10pt "Segoe UI", Arial, sans-serif; color: #000; margin: 0; padding: 16px; }
  h1 { font-size: 16pt; margin: 0 0 2px; }
  .info { color: #605e5c; font-size: 9pt; margin: 0 0 14px; }
  h2 { font-size: 12pt; margin: 18px 0 6px; page-break-after: avoid; break-after: avoid; }
  h2 span { color: #605e5c; font-weight: normal; font-size: 10pt; }
  table { width: 100%; border-collapse: collapse; }
  thead { display: table-header-group; }
  tr { page-break-inside: avoid; break-inside: avoid; }
  th { text-align: left; background: #f3f2f1; border-bottom: 1.5px solid #8a8886; padding: 4px 8px; font-size: 9pt; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  td { padding: 3px 8px; border-bottom: 1px solid #e1dfdd; vertical-align: top; }
  td.num { white-space: nowrap; font-variant-numeric: tabular-nums; }
  .vuoto { color: #605e5c; font-style: italic; }
`;

// Documento HTML completo per la finestra di stampa. Le sezioni senza righe si saltano.
export function generaHtmlStampa(tabelle: ITabella[], info: IInfoStampa): string {
    const conRighe = tabelle.filter(t => t.righe.length > 0);

    const righeInfo = [`Stampata il ${escapeHtml(info.dataStampa)}`];
    if (info.filtro !== "") righeInfo.push(`Filtro: “${escapeHtml(info.filtro)}” (non è l’elenco completo)`);

    const corpo = conRighe.length > 0
        ? conRighe.map(htmlTabella).join("")
        : '<p class="vuoto">Nessun contatto da stampare.</p>';

    return `<!DOCTYPE html><html lang="it"><head><meta charset="utf-8">` +
        `<title>${escapeHtml(info.titolo)}</title><style>${STILE_STAMPA}</style></head>` +
        `<body><h1>${escapeHtml(info.titolo)}</h1><p class="info">${righeInfo.join(" · ")}</p>${corpo}</body></html>`;
}
