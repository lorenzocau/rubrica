import {
    nomeDipendente, costruisciTabelle, campoCsv, generaCsv, generaHtmlStampa,
    formattaDataStampa, nomeFileCsv, escapeHtml, IOpzioniTabelle, ITabella
} from "../tabelle";
import { merge } from "../rubricaLogic";
import { PROPRIETA_DEFAULT, risolviProprieta } from "../proprieta";
import { IRubricaEntry, UtenteGraph } from "../types";

// --- fixture ---
const rossi: UtenteGraph = {
    id: "1", displayName: "Mario Rossi", givenName: "Mario", surname: "Rossi", mail: "mario.rossi@coind.it",
    jobTitle: "Responsabile", mobilePhone: "+39 333 1234567", faxNumber: "512",
    businessPhones: ["+39 051 6543101"], userType: "Member",
};
const deLuca: UtenteGraph = {
    id: "2", displayName: "Anna De Luca", givenName: "Anna", surname: "De Luca", mail: "anna.deluca@coind.it",
    jobTitle: "Impiegata", businessPhones: ["+39 051 6543007"], userType: "Member",
};
const bianchi: UtenteGraph = {
    id: "3", displayName: "Zeno Bianchi", givenName: "Zeno", surname: "Bianchi", mail: "zeno.bianchi@coind.it",
    userType: "Member", // nessun numero: deve comparire lo stesso
};
const ospite: UtenteGraph = {
    id: "4", displayName: "Gianni Ospite", givenName: "Gianni", surname: "Ospite", mail: "gianni@esterno.com",
    mobilePhone: "333 9999999", userType: "Guest",
};
const sala: UtenteGraph = { id: "5", displayName: "Sala Riunioni", jobTitle: "Sala", businessPhones: ["+39 051 6543999"], userType: "Member" };
const senzaCognome: UtenteGraph = {
    id: "6", displayName: "Bruno Verdi", mail: "bruno.verdi@coind.it", businessPhones: ["+39 051 6543222"],
};

const rubrica: IRubricaEntry[] = merge(
    [rossi, deLuca, bianchi, ospite, sala],
    [
        { contatto: "Portineria", interno: "5", reparto: "Logistica" },
        { contatto: "Accettazione", interno: "6" },
    ],
    [
        { nome: "Taxi Navile", numero: "051 7654321", numeroBreve: "301" },
        { nome: "Corriere", numero: "", numeroBreve: "302" },
    ]
);

const tutte: IOpzioniTabelle = { dipendenti: true, servizi: true, esterni: true, mostraCellulare: true, mostraReparto: true };

function tabella(tabelle: ITabella[], id: string): ITabella {
    const t = tabelle.filter(x => x.id === id)[0];
    expect(t).toBeDefined();
    return t;
}

describe("nomeDipendente", () => {
    test("Cognome Nome da surname e givenName (anche cognomi composti)", () => {
        expect(nomeDipendente(rossi)).toBe("Rossi Mario");
        expect(nomeDipendente(deLuca)).toBe("De Luca Anna");
    });

    test("senza cognome si ripiega sul nome visualizzato, poi sulla mail", () => {
        expect(nomeDipendente(senzaCognome)).toBe("Bruno Verdi");
        expect(nomeDipendente({ mail: "x@coind.it" })).toBe("x@coind.it");
        expect(nomeDipendente({})).toBe("");
    });

    test("solo cognome, senza nome", () => {
        expect(nomeDipendente({ surname: " Rossi ", givenName: "  " })).toBe("Rossi");
    });
});

describe("costruisciTabelle: dipendenti", () => {
    const dip = tabella(costruisciTabelle(rubrica, "", tutte), "dipendenti");

    test("ordinati per cognome", () => {
        expect(dip.righe.map(r => r.nome)).toEqual(["Bianchi Zeno", "De Luca Anna", "Rossi Mario"]);
    });

    test("ospiti e account senza mail non compaiono", () => {
        const nomi = dip.righe.map(r => r.nome);
        expect(nomi.indexOf("Ospite Gianni")).toBe(-1);
        expect(nomi.some(n => (n ?? "").indexOf("Sala") !== -1)).toBe(false);
    });

    test("gli utenti senza alcun numero compaiono con le celle vuote", () => {
        const r = dip.righe.filter(x => x.nome === "Bianchi Zeno")[0];
        expect(r.interno).toBe("");
        expect(r.cellulare).toBe("");
        expect(r.cellulareBreve).toBe("");
    });

    test("colonne e valori: interno breve, cellulare, cellulare breve, job title", () => {
        expect(dip.colonne.map(c => c.titolo)).toEqual(
            ["Cognome Nome", "N. interno breve", "Cellulare", "Cellulare breve", "Job title"]
        );
        const r = dip.righe.filter(x => x.nome === "Rossi Mario")[0];
        expect(r).toEqual({
            nome: "Rossi Mario", interno: "101", cellulare: "+39 333 1234567", cellulareBreve: "512", jobTitle: "Responsabile",
        });
        expect(dip.righe.filter(x => x.nome === "De Luca Anna")[0].interno).toBe("007"); // zeri iniziali
    });

    test("mostraCellulare=false toglie colonne e dati dei cellulari", () => {
        const d = tabella(costruisciTabelle(rubrica, "", { ...tutte, mostraCellulare: false }), "dipendenti");
        expect(d.colonne.map(c => c.chiave)).toEqual(["nome", "interno", "jobTitle"]);
        expect(JSON.stringify(d.righe)).not.toContain("1234567");
        expect(JSON.stringify(d.righe)).not.toContain("512");
    });

    test("mostraReparto=false toglie il job title", () => {
        const d = tabella(costruisciTabelle(rubrica, "", { ...tutte, mostraReparto: false }), "dipendenti");
        expect(d.colonne.some(c => c.chiave === "jobTitle")).toBe(false);
        expect(JSON.stringify(d.righe)).not.toContain("Responsabile");
    });

    test("un utente senza cognome è ordinato con il suo nome visualizzato", () => {
        const r = merge([rossi, senzaCognome], []);
        const d = tabella(costruisciTabelle(r, "", tutte), "dipendenti");
        expect(d.righe.map(x => x.nome)).toEqual(["Bruno Verdi", "Rossi Mario"]);
    });
});

describe("costruisciTabelle: servizi ed esterni", () => {
    const tabelle = costruisciTabelle(rubrica, "", tutte);

    test("servizi per nome, con reparto e interno", () => {
        const s = tabella(tabelle, "servizi");
        expect(s.colonne.map(c => c.titolo)).toEqual(["Nome", "Reparto", "Interno"]);
        expect(s.righe).toEqual([
            { nome: "Accettazione", reparto: "", interno: "6" },
            { nome: "Portineria", reparto: "Logistica", interno: "5" },
        ]);
    });

    test("senza mostraReparto la colonna Reparto non c'è", () => {
        const s = tabella(costruisciTabelle(rubrica, "", { ...tutte, mostraReparto: false }), "servizi");
        expect(s.colonne.map(c => c.chiave)).toEqual(["nome", "interno"]);
        expect(JSON.stringify(s.righe)).not.toContain("Logistica");
    });

    test("esterni per nome, con numero breve e numero", () => {
        const e = tabella(tabelle, "esterni");
        expect(e.colonne.map(c => c.titolo)).toEqual(["Nome", "Numero breve", "Numero"]);
        expect(e.righe.map(r => r.nome)).toEqual(["Corriere", "Taxi Navile"]);
        expect(e.righe[1]).toEqual({ nome: "Taxi Navile", interno: "301", cellulare: "051 7654321" });
    });

    test("le sezioni disattivate non vengono costruite", () => {
        const t = costruisciTabelle(rubrica, "", { ...tutte, servizi: false, esterni: false });
        expect(t.map(x => x.id)).toEqual(["dipendenti"]);
    });
});

describe("costruisciTabelle: ricerca", () => {
    test("la ricerca filtra tutte le tabelle", () => {
        const t = costruisciTabelle(rubrica, "taxi", tutte);
        expect(tabella(t, "dipendenti").righe).toHaveLength(0);
        expect(tabella(t, "esterni").righe).toHaveLength(1);
    });

    test("la ricerca non fa riapparire ospiti e account senza mail", () => {
        expect(tabella(costruisciTabelle(rubrica, "ospite", tutte), "dipendenti").righe).toHaveLength(0);
        expect(tabella(costruisciTabelle(rubrica, "sala", tutte), "dipendenti").righe).toHaveLength(0);
    });

    test("spazi soli equivalgono a nessuna ricerca", () => {
        expect(tabella(costruisciTabelle(rubrica, "   ", tutte), "dipendenti").righe).toHaveLength(3);
    });
});

describe("campoCsv", () => {
    test("testo semplice invariato, vuoto resta vuoto", () => {
        expect(campoCsv("Rossi Mario", false)).toBe("Rossi Mario");
        expect(campoCsv("", true)).toBe("");
    });

    test("i numeri diventano testo per Excel (zeri iniziali e prefisso +)", () => {
        expect(campoCsv("007", true)).toBe('"=""007"""');
        expect(campoCsv("+39 333 1234567", true)).toBe('"=""+39 333 1234567"""');
        expect(campoCsv("101 / 102", true)).toBe('"=""101 / 102"""');
    });

    test("un testo che inizia con = + - @ non diventa una formula", () => {
        expect(campoCsv("=SOMMA(A1)", false)).toBe("'=SOMMA(A1)");
        expect(campoCsv("@x", false)).toBe("'@x");
        expect(campoCsv("-1+1", false)).toBe("'-1+1");
        // un "numero" con lettere non è un numero: stessa protezione
        expect(campoCsv("+39 abc", true)).toBe("'+39 abc");
    });

    test("separatore, virgolette e a capo vengono protetti", () => {
        expect(campoCsv("A;B", false)).toBe('"A;B"');
        expect(campoCsv('Dott. "Verdi"', false)).toBe('"Dott. ""Verdi"""');
        expect(campoCsv("riga1\nriga2", false)).toBe('"riga1\nriga2"');
    });
});

describe("generaCsv", () => {
    test("una riga di intestazione e una riga per contatto, con la colonna Tipo", () => {
        const csv = generaCsv(costruisciTabelle(rubrica, "", tutte));
        const righe = csv.split("\r\n");
        expect(righe[righe.length - 1]).toBe(""); // termina con un a capo
        expect(righe[0]).toBe("Tipo;Nome;N. interno breve;Cellulare;Cellulare breve;Job title;Reparto");
        expect(righe).toHaveLength(1 + 3 + 2 + 2 + 1);
        expect(righe[1].indexOf("Dipendente;Bianchi Zeno;")).toBe(0);
        expect(csv).toContain("Servizio;Portineria;");
        expect(csv).toContain("Esterno;Taxi Navile;");
    });

    test("esterni: numero breve sotto 'N. interno breve', numero intero sotto 'Cellulare'", () => {
        const csv = generaCsv(costruisciTabelle(rubrica, "taxi", tutte));
        const riga = csv.split("\r\n")[1];
        expect(riga).toBe('Esterno;Taxi Navile;"=""301""";"=""051 7654321""";;;');
    });

    test("le colonne nascoste non compaiono nel CSV", () => {
        const csv = generaCsv(costruisciTabelle(rubrica, "", { ...tutte, mostraCellulare: false, mostraReparto: false, esterni: false }));
        expect(csv.split("\r\n")[0]).toBe("Tipo;Nome;N. interno breve");
        expect(csv).not.toContain("1234567");
    });

    test("nessuna tabella: solo l'intestazione", () => {
        expect(generaCsv([])).toBe("Tipo\r\n");
    });
});

describe("date", () => {
    const d = new Date(2026, 10, 16, 9, 5); // 16 novembre 2026, 09:05
    test("data di stampa e nome del file", () => {
        expect(formattaDataStampa(d)).toBe("16/11/2026 09:05");
        expect(nomeFileCsv(d)).toBe("rubrica_2026-11-16.csv");
    });
});

describe("generaHtmlStampa", () => {
    const info = { titolo: "Rubrica", dataStampa: "16/11/2026 14:52", filtro: "" };

    test("contiene titolo, data, sezioni, intestazioni e righe", () => {
        const html = generaHtmlStampa(costruisciTabelle(rubrica, "", tutte), info);
        expect(html).toContain("<!DOCTYPE html>");
        expect(html).toContain("<h1>Rubrica</h1>");
        expect(html).toContain("Stampata il 16/11/2026 14:52");
        expect(html).toContain("<h2>Dipendenti <span>(3)</span></h2>");
        expect(html).toContain("<th>Cognome Nome</th>");
        expect(html).toContain("<td>Rossi Mario</td>");
        expect(html).toContain('<td class="num">+39 333 1234567</td>');
        expect(html).not.toContain("Ospite");
    });

    test("indica quando l'elenco è filtrato", () => {
        const html = generaHtmlStampa(costruisciTabelle(rubrica, "taxi", tutte), { ...info, filtro: "taxi" });
        expect(html).toContain("Filtro: “taxi”");
    });

    test("le sezioni vuote non vengono stampate", () => {
        const html = generaHtmlStampa(costruisciTabelle(rubrica, "taxi", tutte), info);
        expect(html).not.toContain("<h2>Dipendenti");
        expect(html).toContain("<h2>Esterni");
    });

    test("nessuna riga: messaggio al posto delle tabelle", () => {
        const html = generaHtmlStampa(costruisciTabelle(rubrica, "zzzzzzz", tutte), info);
        expect(html).toContain("Nessun contatto da stampare.");
        expect(html).not.toContain("<table>");
    });

    test("il testo dei dati e del filtro è protetto (niente HTML iniettato)", () => {
        const strano: UtenteGraph = { id: "x", displayName: "<img src=x onerror=alert(1)>", mail: "x@coind.it", businessPhones: ["+39 051 6543111"] };
        const html = generaHtmlStampa(
            costruisciTabelle(merge([strano], []), "", tutte),
            { ...info, filtro: "<script>" }
        );
        expect(html).not.toContain("<img");
        expect(html).not.toContain("<script>");
        expect(html).toContain("&lt;img src=x onerror=alert(1)&gt;");
    });

    test("escapeHtml", () => {
        expect(escapeHtml(`<a href="x">&'</a>`)).toBe("&lt;a href=&quot;x&quot;&gt;&amp;&#39;&lt;/a&gt;");
    });
});

describe("impostazione abilitaTabelle", () => {
    test("è attiva di default e un valore salvato false resta false", () => {
        expect(PROPRIETA_DEFAULT.abilitaTabelle).toBe(true);
        expect(risolviProprieta({}).abilitaTabelle).toBe(true);
        expect(risolviProprieta({ abilitaTabelle: false }).abilitaTabelle).toBe(false);
    });
});
