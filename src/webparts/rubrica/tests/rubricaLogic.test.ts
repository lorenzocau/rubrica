import { merge, filtra } from "../rubricaLogic";
import { User } from "@microsoft/microsoft-graph-types";
import { IContattoInterno } from "../types";

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