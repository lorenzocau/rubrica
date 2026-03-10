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

});