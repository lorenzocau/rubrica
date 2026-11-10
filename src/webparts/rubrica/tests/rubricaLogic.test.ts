import {
    merge, filtra, applicaVista, ordina, estraiNumeroBreve,
    haTelefono, descrizioneUtente, dividiNumeri, intervalloNumeriBreviEsterni, numeroBreveNonAssegnato,
    normalizzaPageSize, risolviNomeLista, validaNomeLista,
    PAGE_SIZE_DEFAULT, PAGE_SIZE_MIN, PAGE_SIZE_MAX
} from "../rubricaLogic";
import { risolviProprieta, PROPRIETA_DEFAULT } from "../proprieta";
import { User } from "@microsoft/microsoft-graph-types";
import { IContattoEsterno, IContattoInterno, IRubricaEntry, IRubricaWebPartProps, UtenteGraph } from "../types";

// --- fixture riutilizzabili ---
const userConInterno: User = {
    mail: "mario.rossi@coind.it",
    displayName: "Mario Rossi",
    jobTitle: "Qualità - Responsabile",
};

const userSenzaInterno: User = {
    mail: "anna.bianchi@coind.it",
    displayName: "Anna Bianchi",
    jobTitle: "Amministrazione - Impiegata",
};

const userSenzaMail: User = {
    displayName: "Utente Senza Mail",
    jobTitle: "IT",
};

const vociLista: IContattoInterno[] = [
    { contatto: "mario.rossi@coind.it", interno: "101" },
    { contatto: "reception", interno: "0" },
];

// gli esterni non hanno "interno" (hanno numero e numero breve)
function internoDi(entry?: IRubricaEntry): string | null | undefined {
    return entry && entry.tipo !== "esterno" ? entry.interno : undefined;
}

// ============================================================
// merge
// ============================================================
describe("merge", () => {

    test("utente con corrispondenza nella lista riceve il numero interno", () => {
        const result = merge([userConInterno], vociLista);
        const entry = result.find(e => e.tipo === "utente" && e.user.mail === "mario.rossi@coind.it");
        expect(entry).toBeDefined();
        expect(internoDi(entry)).toBe("101");
    });

    test("utente senza corrispondenza viene inserito con interno null", () => {
        const result = merge([userSenzaInterno], vociLista);
        const entry = result.find(e => e.tipo === "utente" && e.user.mail === "anna.bianchi@coind.it");
        expect(entry).toBeDefined();
        expect(internoDi(entry)).toBeNull();
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
        expect(internoDi(entry)).toBeNull();
    });

    test("il matching è case-insensitive", () => {
        const userMailMaiuscola: User = { ...userConInterno, mail: "MARIO.ROSSI@COIND.IT" };
        const result = merge([userMailMaiuscola], vociLista);
        const entry = result.find(e => e.tipo === "utente");
        expect(internoDi(entry)).toBe("101");
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
        expect(internoDi(result[0])).toBe("101");
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
        expect(internoDi(result[0])).toBe("101 / 102");
    });

    test("stesso contatto con lo stesso interno ripetuto non viene duplicato", () => {
        const result = merge([userConInterno], [
            { contatto: "mario.rossi@coind.it", interno: "101" },
            { contatto: "MARIO.ROSSI@coind.it", interno: "101" },
        ]);
        expect(internoDi(result[0])).toBe("101");
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

    test("il reparto della lista finisce nel servizio; senza reparto resta vuoto", () => {
        const result = merge([], [
            { contatto: "Portineria", interno: "5", reparto: "Logistica" },
            { contatto: "Mensa", interno: "6" },
        ]);
        const reparti = result.map(e => e.tipo === "servizio" ? `${e.nome}:${e.reparto}` : "");
        expect(reparti).toEqual(["Mensa:", "Portineria:Logistica"]);
    });

    test("servizio duplicato: si tiene il primo reparto non vuoto", () => {
        const result = merge([], [
            { contatto: "Reception", interno: "0", reparto: "" },
            { contatto: "reception", interno: "1", reparto: " Direzione " },
            { contatto: "RECEPTION", interno: "2", reparto: "Altro" },
        ]);
        expect(result).toHaveLength(1);
        const s = result[0];
        if (s.tipo === "servizio") expect(s.reparto).toBe("Direzione");
    });

    test("il reparto di una voce abbinata a un utente non crea né modifica nulla", () => {
        const result = merge([userConInterno], [{ contatto: "mario.rossi@coind.it", interno: "101", reparto: "Qualità" }]);
        expect(result).toHaveLength(1);
        expect(result[0].tipo).toBe("utente");
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

    test("match su jobTitle", () => {
        const result = filtra("qualità", rubrica);
        expect(result.some(e => e.tipo === "utente" && e.user.jobTitle === "Qualità - Responsabile")).toBe(true);
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
        expect(filtra("qualita", rubrica).some(e => e.tipo === "utente" && e.user.jobTitle === "Qualità - Responsabile")).toBe(true);
        expect(filtra("QUALITÀ", rubrica).some(e => e.tipo === "utente" && e.user.jobTitle === "Qualità - Responsabile")).toBe(true);
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
    const esterni: IContattoEsterno[] = [
        { nome: "Idraulico Verdi", numero: "051 111111", numeroBreve: "300" },
        { nome: "Taxi Navile", numero: "051 222222", numeroBreve: "301" },
    ];

    test("le voci della lista esterni diventano IRubricaEsterno", () => {
        const result = merge([], [], esterni);
        expect(result).toHaveLength(2);
        expect(result.every(e => e.tipo === "esterno")).toBe(true);
    });

    test("numero e numero breve restano distinti", () => {
        const e = merge([], [], esterni)[0];
        expect(e.tipo).toBe("esterno");
        if (e.tipo === "esterno") {
            expect(e.nome).toBe("Idraulico Verdi");
            expect(e.numero).toBe("051 111111");
            expect(e.numeroBreve).toBe("300");
        }
    });

    test("senza il terzo parametro il comportamento resta quello di prima", () => {
        expect(merge([userConInterno], vociLista).some(e => e.tipo === "esterno")).toBe(false);
    });

    test("un esterno non viene mai abbinato a un utente Graph, nemmeno con la stessa mail", () => {
        const result = merge([userConInterno], [], [{ nome: "mario.rossi@coind.it", numero: "999", numeroBreve: "998" }]);
        const utente = result.find(e => e.tipo === "utente");
        expect(internoDi(utente)).toBeNull();
        expect(result.filter(e => e.tipo === "esterno")).toHaveLength(1);
    });

    test("esterno duplicato (maiuscole diverse) viene unito e tutti i numeri sono mostrati", () => {
        const result = merge([], [], [
            { nome: "Taxi Navile", numero: "051 1", numeroBreve: "301" },
            { nome: "taxi navile", numero: "051 2", numeroBreve: "302" },
            { nome: "TAXI NAVILE", numero: "051 1", numeroBreve: "301" },
        ]);
        expect(result).toHaveLength(1);
        const e = result[0];
        if (e.tipo === "esterno") {
            expect(e.nome).toBe("Taxi Navile");
            expect(e.numeroBreve).toBe("301 / 302");
            expect(e.numero).toBe("051 1 / 051 2");
        }
    });

    test("basta uno dei due numeri: l'altro resta vuoto", () => {
        const result = merge([], [], [
            { nome: "Solo breve", numero: "", numeroBreve: "310" },
            { nome: "Solo numero", numero: "051 3", numeroBreve: "  " },
        ]);
        expect(result).toHaveLength(2);
        const [breve, numero] = result.map(e => e.tipo === "esterno" ? [e.numero, e.numeroBreve] : []);
        expect(breve).toEqual(["", "310"]);
        expect(numero).toEqual(["051 3", ""]);
    });

    test("righe incomplete della lista esterni vengono ignorate", () => {
        const result = merge([], [], [
            { nome: " ", numero: "051 1", numeroBreve: "1" },
            { nome: "Corriere", numero: "", numeroBreve: "" },
            { nome: "Corriere", numero: "  ", numeroBreve: " " },
        ]);
        expect(result).toHaveLength(0);
    });

    test("stesso nome in servizi ed esterni: restano due voci distinte", () => {
        const result = merge([], [{ contatto: "Reception", interno: "0" }], [{ nome: "Reception", numero: "", numeroBreve: "9" }]);
        expect(result.map(e => e.tipo).sort()).toEqual(["esterno", "servizio"]);
    });
});

describe("filtra con esterni", () => {
    const rubrica = merge(
        [userConInterno],
        [{ contatto: "Portineria", interno: "5" }],
        [
            { nome: "Idraulico Verdi", numero: "+39 051 1234567", numeroBreve: "300" },
            { nome: "Taxi Navile", numero: "051 7654321", numeroBreve: "301" },
        ]
    );

    test("ricerca per nome sugli esterni", () => {
        const result = filtra("idraulico", rubrica);
        expect(result).toHaveLength(1);
        expect(result[0].tipo).toBe("esterno");
    });

    test("ricerca per numero breve sugli esterni", () => {
        const result = filtra("301", rubrica);
        expect(result).toHaveLength(1);
        expect(result[0].tipo === "esterno" && result[0].nome).toBe("Taxi Navile");
    });

    test("ricerca per numero intero sugli esterni, anche con spazi e prefisso", () => {
        expect(filtra("1234567", rubrica)).toHaveLength(1);
        expect(filtra("051 1234567", rubrica)).toHaveLength(1);
        expect(filtra("+39 051 1234567", rubrica)).toHaveLength(1);
        expect(filtra("0517654321", rubrica)).toHaveLength(1);
        expect(filtra("051 765", rubrica)).toHaveLength(1);
    });

    test("poche cifre non cercano nei numeri interi (altrimenti '123' troverebbe chiunque)", () => {
        expect(filtra("123", rubrica)).toHaveLength(0);
        expect(filtra("051", rubrica)).toHaveLength(0);
    });

    test("nome e numero intero si combinano", () => {
        expect(filtra("taxi 7654321", rubrica)).toHaveLength(1);
        expect(filtra("idraulico 7654321", rubrica)).toHaveLength(0);
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

describe("filtra: reparto dei servizi e numeri dei cellulari", () => {
    const carla: UtenteGraph = {
        id: "c", displayName: "Carla Neri", mail: "carla.neri@coind.it",
        mobilePhone: "+39 333 1234567", faxNumber: "512",
    };
    const rubrica = merge(
        [carla],
        [{ contatto: "Portineria", interno: "5", reparto: "Logistica" }]
    );

    test("il reparto dei servizi è ricercabile", () => {
        const result = filtra("logistica", rubrica);
        expect(result).toHaveLength(1);
        expect(result[0].tipo).toBe("servizio");
    });

    test("il numero breve del cellulare (faxNumber) è ricercabile", () => {
        expect(filtra("512", rubrica)).toHaveLength(1);
    });

    test("il cellulare intero è ricercabile in vari formati", () => {
        expect(filtra("3331234567", rubrica)).toHaveLength(1);
        expect(filtra("333 1234567", rubrica)).toHaveLength(1);
        expect(filtra("+39 333 1234567", rubrica)).toHaveLength(1);
        expect(filtra("1234", rubrica)).toHaveLength(1);
        expect(filtra("carla 1234", rubrica)).toHaveLength(1);
    });

    test("poche cifre non trovano nel cellulare intero", () => {
        expect(filtra("123", rubrica)).toHaveLength(0);
        expect(filtra("333", rubrica)).toHaveLength(0);
    });

    test("cifre che non corrispondono non trovano nulla", () => {
        expect(filtra("9999", rubrica)).toHaveLength(0);
        expect(filtra("carla 9999", rubrica)).toHaveLength(0);
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
            [{ nome: "Corriere", numero: "", numeroBreve: "7" }]
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
        [{ nome: "Idraulico", numero: "051 1", numeroBreve: "300" }]
    );

    const tutto = { soloMembri: false, nascondiSenzaMail: false, nascondiSenzaTelefono: false, includiEsterni: true };

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
        const r = merge([], [{ contatto: "Portineria", interno: "5" }], [{ nome: "Taxi", numero: "", numeroBreve: "6" }]);
        expect(applicaVista(r, { soloMembri: true, nascondiSenzaMail: true, nascondiSenzaTelefono: true, includiEsterni: true })).toHaveLength(2);
    });
});

// ============================================================
// utenti senza telefono
// ============================================================
describe("nascondiSenzaTelefono", () => {
    const senzaNumeri: UtenteGraph = { id: "n1", displayName: "Senza Numeri", mail: "senza@coind.it" };
    const telefonoCorto: UtenteGraph = { id: "n2", displayName: "Telefono Corto", mail: "corto@coind.it", businessPhones: ["12"] };
    const soloFisso: UtenteGraph = { id: "n3", displayName: "Solo Fisso", mail: "fisso@coind.it", businessPhones: ["+39 051 6543123"] };
    const soloCell: UtenteGraph = { id: "n4", displayName: "Solo Cell", mail: "cell@coind.it", mobilePhone: "+39 333 1234567" };
    const soloBreve: UtenteGraph = { id: "n5", displayName: "Solo Breve", mail: "breve@coind.it", faxNumber: "512" };
    const spazi: UtenteGraph = { id: "n6", displayName: "Solo Spazi", mail: "spazi@coind.it", mobilePhone: "  ", faxNumber: "" };
    const conVoce: UtenteGraph = { id: "n7", displayName: "Con Voce", mail: "voce@coind.it" };

    const rubrica = merge(
        [senzaNumeri, telefonoCorto, soloFisso, soloCell, soloBreve, spazi, conVoce],
        [{ contatto: "voce@coind.it", interno: "77" }, { contatto: "Portineria", interno: "5" }],
        [{ nome: "Taxi", numero: "", numeroBreve: "6" }]
    );
    const opz = { soloMembri: false, nascondiSenzaMail: false, nascondiSenzaTelefono: true, includiEsterni: true };
    const nomi = (r: IRubricaEntry[]) => r.map(e => e.tipo === "utente" ? e.user.displayName : e.nome);

    test("con l'opzione attiva nasconde solo gli utenti senza nessun numero", () => {
        expect(nomi(applicaVista(rubrica, opz))).toEqual(
            ["Con Voce", "Portineria", "Solo Breve", "Solo Cell", "Solo Fisso", "Taxi"]
        );
    });

    test("con l'opzione spenta non nasconde nulla", () => {
        expect(applicaVista(rubrica, { ...opz, nascondiSenzaTelefono: false })).toHaveLength(rubrica.length);
    });

    test("un interno preso dalla lista conta come numero", () => {
        expect(nomi(applicaVista(rubrica, opz))).toContain("Con Voce");
    });

    test("un telefono fisso troppo corto (nessun interno ricavabile) non conta", () => {
        expect(nomi(applicaVista(rubrica, opz))).not.toContain("Telefono Corto");
    });

    test("cellulare e numero breve fatti solo di spazi non contano", () => {
        expect(nomi(applicaVista(rubrica, opz))).not.toContain("Solo Spazi");
    });

    test("servizi ed esterni non sono toccati", () => {
        const r = applicaVista(rubrica, opz);
        expect(r.some(e => e.tipo === "servizio")).toBe(true);
        expect(r.some(e => e.tipo === "esterno")).toBe(true);
    });

    test("haTelefono guarda interno, cellulare e numero breve", () => {
        const u = (user: UtenteGraph, interno: string | null = null) => ({ tipo: "utente" as const, user, interno });
        expect(haTelefono(u(senzaNumeri))).toBe(false);
        expect(haTelefono(u(senzaNumeri, "101"))).toBe(true);
        expect(haTelefono(u(soloCell))).toBe(true);
        expect(haTelefono(u(soloBreve))).toBe(true);
        expect(haTelefono(u(senzaNumeri, "  "))).toBe(false);
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

    test("nascondiSenzaTelefono è attivo di default, ma un valore salvato false resta false", () => {
        expect(risolviProprieta({}).nascondiSenzaTelefono).toBe(true);
        expect(risolviProprieta({ nascondiSenzaTelefono: false }).nascondiSenzaTelefono).toBe(false);
    });

    test("non modifica l'oggetto dei default", () => {
        risolviProprieta({ risultatiPerPagina: 99 });
        expect(PROPRIETA_DEFAULT.risultatiPerPagina).toBe(30);
    });
});


// ============================================================
// numero breve (ultime 3 cifre del telefono in AD)
// ============================================================
describe("estraiNumeroBreve", () => {
    test("prende le ultime 3 cifre", () => {
        expect(estraiNumeroBreve("+39 051 6543123")).toBe("123");
        expect(estraiNumeroBreve("0516543123")).toBe("123");
    });

    test("ignora spazi, trattini e altri separatori", () => {
        expect(estraiNumeroBreve("051-654 31.23")).toBe("123");
        expect(estraiNumeroBreve("(051) 654 3 123 ")).toBe("123");
    });

    test("mantiene gli zeri iniziali", () => {
        expect(estraiNumeroBreve("+39 051 6543007")).toBe("007");
    });

    test("un numero di sole 3 cifre vale come numero breve", () => {
        expect(estraiNumeroBreve("123")).toBe("123");
    });

    test("meno di 3 cifre, vuoto o assente: null", () => {
        expect(estraiNumeroBreve("12")).toBeNull();
        expect(estraiNumeroBreve("")).toBeNull();
        expect(estraiNumeroBreve("   ")).toBeNull();
        expect(estraiNumeroBreve("abc")).toBeNull();
        expect(estraiNumeroBreve(undefined)).toBeNull();
        expect(estraiNumeroBreve(null)).toBeNull();
    });
});

describe("merge con numero breve", () => {
    const conTelefono: User = { id: "t", displayName: "Luca Verdi", mail: "luca.verdi@coind.it", businessPhones: ["+39 051 6543456"] };

    test("l'utente senza voce in lista usa il numero breve del telefono", () => {
        const result = merge([conTelefono], []);
        expect(internoDi(result[0])).toBe("456");
    });

    test("una voce in lista con la sua mail sostituisce il numero breve", () => {
        const result = merge([conTelefono], [{ contatto: "luca.verdi@coind.it", interno: "900" }]);
        expect(internoDi(result[0])).toBe("900");
    });

    test("si usa solo il primo numero di businessPhones", () => {
        const result = merge([{ ...conTelefono, businessPhones: ["+39 051 6543111", "+39 051 6543222"] }], []);
        expect(internoDi(result[0])).toBe("111");
    });

    test("senza telefono (array vuoto o assente) l'interno resta null", () => {
        expect(internoDi(merge([{ ...conTelefono, businessPhones: [] }], [])[0])).toBeNull();
        expect(internoDi(merge([{ ...conTelefono, businessPhones: undefined }], [])[0])).toBeNull();
    });

    test("la ricerca trova l'utente dal numero breve", () => {
        const rubrica = merge([conTelefono], []);
        expect(filtra("456", rubrica)).toHaveLength(1);
    });
});


// ============================================================
// numeri brevi liberi (ricerca di un numero non assegnato)
// ============================================================
describe("descrizioneUtente", () => {
    test("reparto e job title separati da un trattino, prima il reparto", () => {
        expect(descrizioneUtente({ department: "IT", jobTitle: "Sistemista" })).toBe("IT - Sistemista");
    });

    test("se ne manca uno si mostra solo l'altro, senza trattini in più", () => {
        expect(descrizioneUtente({ department: "IT" })).toBe("IT");
        expect(descrizioneUtente({ jobTitle: " Sistemista " })).toBe("Sistemista");
        expect(descrizioneUtente({ department: "  ", jobTitle: null })).toBe("");
        expect(descrizioneUtente({})).toBe("");
    });

    test("la ricerca trova sia per reparto sia per job title", () => {
        const r = merge([{ id: "x", displayName: "Eva Rossi", mail: "eva@coind.it", department: "Logistica", jobTitle: "Spedizioniere" }], []);
        expect(filtra("logistica", r)).toHaveLength(1);
        expect(filtra("spedizioniere", r)).toHaveLength(1);
        expect(filtra("logistica spedizioniere", r)).toHaveLength(1);
    });
});

describe("dividiNumeri", () => {
    test("separa i numeri uniti con ' / ' e scarta i vuoti", () => {
        expect(dividiNumeri("101 / 102")).toEqual(["101", "102"]);
        expect(dividiNumeri("101")).toEqual(["101"]);
        expect(dividiNumeri("")).toEqual([]);
        expect(dividiNumeri(undefined)).toEqual([]);
    });

    test("una barra senza spazi fa parte del numero", () => {
        expect(dividiNumeri("051/123456")).toEqual(["051/123456"]);
    });
});

describe("numeroBreveNonAssegnato", () => {
    const ugo: UtenteGraph = {
        id: "u", displayName: "Ugo Verdi", mail: "ugo@coind.it",
        businessPhones: ["+39 051 6543123"], faxNumber: "200",
    };
    const rubrica = merge(
        [ugo],
        [{ contatto: "Portineria", interno: "150" }],
        [
            { nome: "Primo", numero: "", numeroBreve: "100" },
            { nome: "Secondo", numero: "051 2", numeroBreve: "301 / 302" },
            { nome: "Ultimo", numero: "", numeroBreve: "551" },
            { nome: "Senza breve", numero: "051 3", numeroBreve: "" },
        ]
    );

    test("l'intervallo si ricava dai numeri brevi degli esterni", () => {
        expect(intervalloNumeriBreviEsterni(rubrica)).toEqual({ min: 100, max: 551 });
    });

    test("l'intervallo è dinamico: segue i dati", () => {
        const r = merge([], [], [
            { nome: "A", numero: "", numeroBreve: "10" },
            { nome: "B", numero: "", numeroBreve: "20" },
        ]);
        expect(intervalloNumeriBreviEsterni(r)).toEqual({ min: 10, max: 20 });
        expect(numeroBreveNonAssegnato("15", r)).toEqual({ numero: "15", min: 10, max: 20 });
        expect(numeroBreveNonAssegnato("21", r)).toBeNull();
    });

    test("un numero dentro l'intervallo e senza nessuna voce è segnalato", () => {
        expect(numeroBreveNonAssegnato("250", rubrica)).toEqual({ numero: "250", min: 100, max: 551 });
        expect(numeroBreveNonAssegnato("  250 ", rubrica)).toEqual({ numero: "250", min: 100, max: 551 });
    });

    test("estremi dell'intervallo: assegnati (esterni) quindi non segnalati", () => {
        expect(numeroBreveNonAssegnato("100", rubrica)).toBeNull();
        expect(numeroBreveNonAssegnato("551", rubrica)).toBeNull();
    });

    test("un numero assegnato a qualunque voce non è segnalato", () => {
        expect(numeroBreveNonAssegnato("302", rubrica)).toBeNull(); // secondo numero di un esterno
        expect(numeroBreveNonAssegnato("123", rubrica)).toBeNull(); // interno ricavato dal telefono dell'utente
        expect(numeroBreveNonAssegnato("200", rubrica)).toBeNull(); // numero breve del cellulare (faxNumber)
        expect(numeroBreveNonAssegnato("150", rubrica)).toBeNull(); // interno di un servizio
    });

    test("fuori dall'intervallo non è segnalato", () => {
        expect(numeroBreveNonAssegnato("99", rubrica)).toBeNull();
        expect(numeroBreveNonAssegnato("552", rubrica)).toBeNull();
    });

    test("testo o numeri con separatori non sono numeri brevi", () => {
        expect(numeroBreveNonAssegnato("abc", rubrica)).toBeNull();
        expect(numeroBreveNonAssegnato("2 50", rubrica)).toBeNull();
        expect(numeroBreveNonAssegnato("+250", rubrica)).toBeNull();
        expect(numeroBreveNonAssegnato("", rubrica)).toBeNull();
    });

    test("senza esterni con numero breve non c'è nessun intervallo", () => {
        expect(intervalloNumeriBreviEsterni(merge([ugo], []))).toBeNull();
        expect(numeroBreveNonAssegnato("250", merge([ugo], []))).toBeNull();
        expect(numeroBreveNonAssegnato("250", merge([], [], [{ nome: "X", numero: "051 1", numeroBreve: "" }]))).toBeNull();
    });
});
