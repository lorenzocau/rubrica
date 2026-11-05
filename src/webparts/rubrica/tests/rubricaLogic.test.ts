import {
    merge, filtra, applicaVista, ordina,
    normalizzaPageSize, risolviNomeLista, validaNomeLista,
    PAGE_SIZE_DEFAULT, PAGE_SIZE_MIN, PAGE_SIZE_MAX
} from "../rubricaLogic";
import { risolviProprieta, PROPRIETA_DEFAULT } from "../proprieta";
import { User } from "@microsoft/microsoft-graph-types";
import { IContattoInterno, IRubricaWebPartProps } from "../types";

// --- fixture riutilizzabili ---
const userConInterno: User = {
    mail: "mario.rossi@coind.it",
    displayName: "Mario Rossi",
    department: "Qualità",
    jobTitle: "Responsabile",
};

const userSenzaInterno: User = {
    mail: "anna.bianchi@coind.it",
    displayName: "Anna Bianchi",
    department: "Amministrazione",
    jobTitle: "Impiegata",
};

const userSenzaMail: User = {
    displayName: "Utente Senza Mail",
    department: "IT",
};

const vociLista: IContattoInterno[] = [
    { contatto: "mario.rossi@coind.it", interno: "101" },
    { contatto: "reception", interno: "0" },
];

// ============================================================
// merge
// ============================================================
describe("merge", () => {

    test("utente con corrispondenza nella lista riceve il numero interno", () => {
        const result = merge([userConInterno], vociLista);
        const entry = result.find(e => e.tipo === "utente" && e.user.mail === "mario.rossi@coind.it");
        expect(entry).toBeDefined();
        expect(entry?.interno).toBe("101");
    });

    test("utente senza corrispondenza viene inserito con interno null", () => {
        const result = merge([userSenzaInterno], vociLista);
        const entry = result.find(e => e.tipo === "utente" && e.user.mail === "anna.bianchi@coind.it");
        expect(entry).toBeDefined();
        expect(entry?.interno).toBeNull();
    });

    test("voce lista senza corrispondenza in Graph diventa IRubricaServizio", () => {
        const result = merge([userConInterno], vociLista);
        const servizio = result.find(e => e.tipo === "servizio");
        expect(servizio).toBeDefined();
        if (servizio?.tipo === "servizio") {
            expect(servizio.nome).toBe("reception");
            expect(servizio.interno).toBe("0");
        }
    });

    test("utente senza email viene inserito con interno null senza errori", () => {
        const result = merge([userSenzaMail], vociLista);
        const entry = result.find(e => e.tipo === "utente");
        expect(entry).toBeDefined();
        expect(entry?.interno).toBeNull();
    });

    test("il matching è case-insensitive", () => {
        const userMailMaiuscola: User = { ...userConInterno, mail: "MARIO.ROSSI@COIND.IT" };
        const result = merge([userMailMaiuscola], vociLista);
        const entry = result.find(e => e.tipo === "utente");
        expect(entry?.interno).toBe("101");
    });

    test("entrambi gli array vuoti restituiscono array vuoto", () => {
        expect(merge([], [])).toHaveLength(0);
    });

    test("il nome del servizio mantiene le maiuscole originali", () => {
        const result = merge([], [{ contatto: "Magazzino Materie Prime", interno: "200" }]);
        const servizio = result[0];
        expect(servizio.tipo).toBe("servizio");
        if (servizio.tipo === "servizio") {
            expect(servizio.nome).toBe("Magazzino Materie Prime");
        }
    });

    test("spazi ai bordi di contatto e interno vengono ignorati", () => {
        const result = merge([userConInterno], [{ contatto: "  Mario.Rossi@coind.it ", interno: " 101 " }]);
        expect(result).toHaveLength(1); // nessun servizio orfano
        expect(result[0].interno).toBe("101");
    });

    test("righe con contatto o interno vuoti vengono ignorate", () => {
        const result = merge([], [
            { contatto: "   ", interno: "1" },
            { contatto: "Portineria", interno: " " },
            { contatto: "", interno: "" },
        ]);
        expect(result).toHaveLength(0);
    });

    test("stesso contatto con più interni: vengono mostrati tutti", () => {
        const result = merge([userConInterno], [
            { contatto: "mario.rossi@coind.it", interno: "101" },
            { contatto: "mario.rossi@coind.it", interno: "102" },
        ]);
        expect(result[0].interno).toBe("101 / 102");
    });

    test("stesso contatto con lo stesso interno ripetuto non viene duplicato", () => {
        const result = merge([userConInterno], [
            { contatto: "mario.rossi@coind.it", interno: "101" },
            { contatto: "MARIO.ROSSI@coind.it", interno: "101" },
        ]);
        expect(result[0].interno).toBe("101");
    });

    test("servizio duplicato (maiuscole diverse) viene unito in una sola voce", () => {
        const result = merge([], [
            { contatto: "Reception", interno: "0" },
            { contatto: "reception", interno: "1" },
        ]);
        expect(result).toHaveLength(1);
        const servizio = result[0];
        if (servizio.tipo === "servizio") {
            expect(servizio.nome).toBe("Reception");
            expect(servizio.interno).toBe("0 / 1");
        }
    });

    test("voce con una mail senza utente in Graph diventa un servizio", () => {
        const result = merge([], [{ contatto: "mario.rossi@coind.it", interno: "101" }]);
        expect(result).toHaveLength(1);
        expect(result[0].tipo).toBe("servizio");
    });

    test("due utenti con la stessa mail ricevono entrambi l'interno", () => {
        const result = merge([userConInterno, { ...userConInterno, id: "altro" }], vociLista);
        const utenti = result.filter(e => e.tipo === "utente");
        expect(utenti).toHaveLength(2);
        expect(utenti[0].interno).toBe("101");
        expect(utenti[1].interno).toBe("101");
    });

});

// ============================================================
// filtra
// ============================================================
describe("filtra", () => {
    const rubrica = merge([userConInterno, userSenzaInterno], vociLista);

    test("stringa vuota restituisce sempre array vuoto", () => {
        expect(filtra("", rubrica)).toHaveLength(0);
        expect(filtra("   ", rubrica)).toHaveLength(0);
    });

    test("match su displayName", () => {
        const result = filtra("mario", rubrica);
        expect(result.length).toBeGreaterThan(0);
        expect(result[0].tipo).toBe("utente");
    });

    test("match su department", () => {
        const result = filtra("qualità", rubrica);
        expect(result.some(e => e.tipo === "utente" && e.user.department === "Qualità")).toBe(true);
    });

    test("match su interno", () => {
        const result = filtra("101", rubrica);
        expect(result.some(e => e.tipo === "utente" && e.interno === "101")).toBe(true);
    });

    test("nessun match restituisce array vuoto", () => {
        expect(filtra("zzzzzzz", rubrica)).toHaveLength(0);
    });

    test("servizio trovato per nome", () => {
        const result = filtra("reception", rubrica);
        expect(result.some(e => e.tipo === "servizio")).toBe(true);
    });

    test("la ricerca ignora gli accenti", () => {
        expect(filtra("qualita", rubrica).some(e => e.tipo === "utente" && e.user.department === "Qualità")).toBe(true);
        expect(filtra("QUALITÀ", rubrica).some(e => e.tipo === "utente" && e.user.department === "Qualità")).toBe(true);
    });

    test("più parole in ordine diverso trovano lo stesso utente", () => {
        const result = filtra("rossi mario", rubrica);
        expect(result).toHaveLength(1);
        expect(result[0].tipo).toBe("utente");
    });

    test("le parole possono corrispondere a campi diversi", () => {
        const result = filtra("mario qualità", rubrica);
        expect(result).toHaveLength(1);
    });

    test("tutte le parole devono corrispondere", () => {
        expect(filtra("mario amministrazione", rubrica)).toHaveLength(0);
    });

    test("una singola lettera non restituisce tutti i servizi", () => {
        const servizi = merge([], [
            { contatto: "Portineria", interno: "5" },
            { contatto: "Magazzino", interno: "6" },
        ]);
        const result = filtra("e", servizi);
        expect(result).toHaveLength(1); // solo "Portineria" contiene la e
    });

    test("la parola 'servizio' trova tutti i servizi", () => {
        const servizi = merge([], [
            { contatto: "Portineria", interno: "5" },
            { contatto: "Magazzino", interno: "6" },
        ]);
        expect(filtra("servizio", servizi)).toHaveLength(2);
        expect(filtra("servizio magazzino", servizi)).toHaveLength(1);
    });

    test("la parola 'servizio' non trova gli utenti", () => {
        expect(filtra("servizio", merge([userConInterno], []))).toHaveLength(0);
    });

    test("ricerca su uno degli interni multipli", () => {
        const multi = merge([userConInterno], [
            { contatto: "mario.rossi@coind.it", interno: "101" },
            { contatto: "mario.rossi@coind.it", interno: "102" },
        ]);
        expect(filtra("102", multi)).toHaveLength(1);
    });

});

// ============================================================
// esterni
// ============================================================
describe("merge con esterni", () => {
    const esterni: IContattoInterno[] = [
        { contatto: "Idraulico Verdi", interno: "300" },
        { contatto: "Taxi Navile", interno: "301" },
    ];

    test("le voci della lista esterni diventano IRubricaEsterno", () => {
        const result = merge([], [], esterni);
        expect(result).toHaveLength(2);
        expect(result.every(e => e.tipo === "esterno")).toBe(true);
    });

    test("senza il terzo parametro il comportamento resta quello di prima", () => {
        expect(merge([userConInterno], vociLista).some(e => e.tipo === "esterno")).toBe(false);
    });

    test("un esterno non viene mai abbinato a un utente Graph, nemmeno con la stessa mail", () => {
        const result = merge([userConInterno], [], [{ contatto: "mario.rossi@coind.it", interno: "999" }]);
        const utente = result.find(e => e.tipo === "utente");
        expect(utente?.interno).toBeNull();
        expect(result.filter(e => e.tipo === "esterno")).toHaveLength(1);
    });

    test("esterno duplicato (maiuscole diverse) viene unito e gli interni sono tutti mostrati", () => {
        const result = merge([], [], [
            { contatto: "Taxi Navile", interno: "301" },
            { contatto: "taxi navile", interno: "302" },
            { contatto: "TAXI NAVILE", interno: "301" },
        ]);
        expect(result).toHaveLength(1);
        const e = result[0];
        if (e.tipo === "esterno") {
            expect(e.nome).toBe("Taxi Navile");
            expect(e.interno).toBe("301 / 302");
        }
    });

    test("righe incomplete della lista esterni vengono ignorate", () => {
        const result = merge([], [], [
            { contatto: " ", interno: "1" },
            { contatto: "Corriere", interno: "" },
        ]);
        expect(result).toHaveLength(0);
    });

    test("stesso nome in servizi ed esterni: restano due voci distinte", () => {
        const result = merge([], [{ contatto: "Reception", interno: "0" }], [{ contatto: "Reception", interno: "9" }]);
        expect(result.map(e => e.tipo).sort()).toEqual(["esterno", "servizio"]);
    });
});

describe("filtra con esterni", () => {
    const rubrica = merge(
        [userConInterno],
        [{ contatto: "Portineria", interno: "5" }],
        [{ contatto: "Idraulico Verdi", interno: "300" }, { contatto: "Taxi Navile", interno: "301" }]
    );

    test("ricerca per nome sugli esterni", () => {
        const result = filtra("idraulico", rubrica);
        expect(result).toHaveLength(1);
        expect(result[0].tipo).toBe("esterno");
    });

    test("ricerca per interno sugli esterni", () => {
        expect(filtra("301", rubrica)).toHaveLength(1);
    });

    test("la parola 'esterno' (o 'esterni') trova tutti gli esterni e solo quelli", () => {
        const a = filtra("esterno", rubrica);
        const b = filtra("esterni", rubrica);
        expect(a).toHaveLength(2);
        expect(b).toHaveLength(2);
        expect(a.every(e => e.tipo === "esterno")).toBe(true);
    });

    test("'esterno' si combina con altre parole", () => {
        expect(filtra("esterno taxi", rubrica)).toHaveLength(1);
    });

    test("la parola 'servizio' non trova gli esterni", () => {
        const result = filtra("servizio", rubrica);
        expect(result).toHaveLength(1);
        expect(result[0].tipo).toBe("servizio");
    });
});

// ============================================================
// ordinamento
// ============================================================
describe("ordinamento", () => {
    test("merge restituisce utenti, servizi ed esterni in ordine alfabetico", () => {
        const result = merge(
            [{ displayName: "Zeno Zanni", mail: "zeno@coind.it" }, { displayName: "Anna Bianchi", mail: "anna@coind.it" }],
            [{ contatto: "Magazzino", interno: "6" }],
            [{ contatto: "Corriere", interno: "7" }]
        );
        const nomi = result.map(e => e.tipo === "utente" ? e.user.displayName : e.nome);
        expect(nomi).toEqual(["Anna Bianchi", "Corriere", "Magazzino", "Zeno Zanni"]);
    });

    test("l'ordine ignora maiuscole e accenti", () => {
        const result = merge(
            [{ displayName: "bruno", mail: "b@coind.it" }, { displayName: "Álvaro", mail: "a@coind.it" }, { displayName: "Carlo", mail: "c@coind.it" }],
            []
        );
        const nomi = result.map(e => e.tipo === "utente" ? e.user.displayName : e.nome);
        expect(nomi).toEqual(["Álvaro", "bruno", "Carlo"]);
    });

    test("a parità di nome l'ordine di arrivo viene mantenuto", () => {
        const result = ordina([
            { tipo: "utente", user: { id: "1", displayName: "Mario Rossi" }, interno: "1" },
            { tipo: "utente", user: { id: "2", displayName: "Mario Rossi" }, interno: "2" },
        ]);
        expect(result.map(e => e.tipo === "utente" ? e.user.id : "")).toEqual(["1", "2"]);
    });

    test("un utente senza displayName viene ordinato per mail", () => {
        const result = merge([{ mail: "bbb@coind.it" }, { displayName: "Aaa", mail: "x@coind.it" }], []);
        expect(result[0].tipo === "utente" && result[0].user.displayName).toBe("Aaa");
    });
});

// ============================================================
// applicaVista
// ============================================================
describe("applicaVista", () => {
    const ospite: User = { id: "g", displayName: "Gianni Ospite", mail: "gianni@esterno.com", userType: "Guest" };
    const sala: User = { id: "s", displayName: "Sala Riunioni", userType: "Member" };
    const rubrica = merge(
        [userConInterno, ospite, sala],
        [{ contatto: "gianni@esterno.com", interno: "77" }],
        [{ contatto: "Idraulico", interno: "300" }]
    );

    const tutto = { soloMembri: false, nascondiSenzaMail: false, includiEsterni: true };

    test("senza filtri non toglie nulla", () => {
        expect(applicaVista(rubrica, tutto)).toHaveLength(rubrica.length);
    });

    test("includiEsterni=false toglie solo gli esterni", () => {
        const result = applicaVista(rubrica, { ...tutto, includiEsterni: false });
        expect(result.some(e => e.tipo === "esterno")).toBe(false);
        expect(result).toHaveLength(rubrica.length - 1);
    });

    test("soloMembri nasconde gli ospiti Guest ma non gli altri", () => {
        const result = applicaVista(rubrica, { ...tutto, soloMembri: true });
        expect(result.some(e => e.tipo === "utente" && e.user.userType === "Guest")).toBe(false);
        expect(result.some(e => e.tipo === "utente" && e.user.id === "s")).toBe(true);
    });

    test("un utente senza userType (null/undefined) non viene scambiato per ospite", () => {
        const result = applicaVista(merge([userConInterno], []), { ...tutto, soloMembri: true });
        expect(result).toHaveLength(1);
    });

    test("nascondere un ospite NON fa ricomparire la sua voce di lista come servizio", () => {
        const result = applicaVista(rubrica, { ...tutto, soloMembri: true });
        expect(result.some(e => e.tipo === "servizio")).toBe(false);
    });

    test("nascondiSenzaMail toglie gli account senza mail (anche con mail vuota)", () => {
        const conMailVuota: User = { id: "v", displayName: "Vuota", mail: "  " };
        const r = merge([userConInterno, sala, conMailVuota], []);
        const result = applicaVista(r, { ...tutto, nascondiSenzaMail: true });
        expect(result).toHaveLength(1);
        expect(result[0].tipo === "utente" && result[0].user.mail).toBe("mario.rossi@coind.it");
    });

    test("i filtri sugli utenti non toccano servizi ed esterni", () => {
        const r = merge([], [{ contatto: "Portineria", interno: "5" }], [{ contatto: "Taxi", interno: "6" }]);
        expect(applicaVista(r, { soloMembri: true, nascondiSenzaMail: true, includiEsterni: true })).toHaveLength(2);
    });
});

// ============================================================
// impostazioni
// ============================================================
describe("normalizzaPageSize", () => {
    test("valori validi restano invariati", () => {
        expect(normalizzaPageSize(30)).toBe(30);
        expect(normalizzaPageSize(PAGE_SIZE_MIN)).toBe(PAGE_SIZE_MIN);
        expect(normalizzaPageSize(PAGE_SIZE_MAX)).toBe(PAGE_SIZE_MAX);
    });

    test("valori fuori limite vengono riportati nei limiti", () => {
        expect(normalizzaPageSize(1)).toBe(PAGE_SIZE_MIN);
        expect(normalizzaPageSize(0)).toBe(PAGE_SIZE_MIN);
        expect(normalizzaPageSize(-10)).toBe(PAGE_SIZE_MIN);
        expect(normalizzaPageSize(5000)).toBe(PAGE_SIZE_MAX);
    });

    test("valori non numerici usano il default", () => {
        expect(normalizzaPageSize(NaN)).toBe(PAGE_SIZE_DEFAULT);
        expect(normalizzaPageSize(undefined)).toBe(PAGE_SIZE_DEFAULT);
        expect(normalizzaPageSize(null)).toBe(PAGE_SIZE_DEFAULT);
    });

    test("i decimali vengono arrotondati", () => {
        expect(normalizzaPageSize(12.6)).toBe(13);
    });
});

describe("risolviNomeLista", () => {
    test("campo vuoto o solo spazi usa il default", () => {
        expect(risolviNomeLista("", "Esterni")).toBe("Esterni");
        expect(risolviNomeLista("   ", "Esterni")).toBe("Esterni");
        expect(risolviNomeLista(undefined, "Esterni")).toBe("Esterni");
        expect(risolviNomeLista(null, "Esterni")).toBe("Esterni");
    });

    test("un valore presente viene usato senza spazi ai bordi", () => {
        expect(risolviNomeLista("  Fornitori ", "Esterni")).toBe("Fornitori");
        expect(risolviNomeLista("/sites/intranet/Lists/Esterni", "Esterni")).toBe("/sites/intranet/Lists/Esterni");
    });
});

describe("validaNomeLista", () => {
    test("vuoto e nomi normali sono validi", () => {
        expect(validaNomeLista("")).toBe("");
        expect(validaNomeLista(undefined)).toBe("");
        expect(validaNomeLista("Esterni")).toBe("");
        expect(validaNomeLista("/sites/intranet/Lists/Esterni")).toBe("");
    });

    test("caratteri non validi e nomi troppo lunghi sono rifiutati", () => {
        expect(validaNomeLista("Lista<script>")).not.toBe("");
        expect(validaNomeLista("a".repeat(256))).not.toBe("");
    });
});

describe("risolviProprieta", () => {
    test("senza proprietà salvate restituisce i default", () => {
        expect(risolviProprieta(undefined)).toEqual(PROPRIETA_DEFAULT);
        expect(risolviProprieta(null)).toEqual(PROPRIETA_DEFAULT);
        expect(risolviProprieta({})).toEqual(PROPRIETA_DEFAULT);
    });

    test("i valori salvati hanno la precedenza, anche false e stringa vuota", () => {
        const r = risolviProprieta({ risultatiPerPagina: 12, mostraEsterni: false, titolo: "", listaEsterni: "Fornitori" });
        expect(r.risultatiPerPagina).toBe(12);
        expect(r.mostraEsterni).toBe(false);
        expect(r.listaEsterni).toBe("Fornitori");
        expect(r.listaServizi).toBe(PROPRIETA_DEFAULT.listaServizi);
    });

    test("un web part già in pagina con le vecchie proprietà (solo 'description') ottiene i default", () => {
        const vecchie = { description: "rubrica" } as unknown as Partial<IRubricaWebPartProps>;
        expect(risolviProprieta(vecchie)).toEqual(PROPRIETA_DEFAULT);
    });

    test("valori del tipo sbagliato vengono scartati", () => {
        const sporche = { risultatiPerPagina: "50", mostraMail: "no" } as unknown as Partial<IRubricaWebPartProps>;
        const r = risolviProprieta(sporche);
        expect(r.risultatiPerPagina).toBe(PROPRIETA_DEFAULT.risultatiPerPagina);
        expect(r.mostraMail).toBe(PROPRIETA_DEFAULT.mostraMail);
    });

    test("non modifica l'oggetto dei default", () => {
        risolviProprieta({ risultatiPerPagina: 99 });
        expect(PROPRIETA_DEFAULT.risultatiPerPagina).toBe(30);
    });
});
