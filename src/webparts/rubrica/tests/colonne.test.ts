import { normalizzaNomeColonna, risolviColonne, testoCampo, ICampoLista, IRichiestaColonna } from "../colonne";

describe("normalizzaNomeColonna", () => {
    test("nome interno e nome visualizzato diventano uguali", () => {
        expect(normalizzaNomeColonna("Numero_x0020_breve")).toBe("numerobreve");
        expect(normalizzaNomeColonna("Numero breve")).toBe("numerobreve");
        expect(normalizzaNomeColonna("NumeroBreve")).toBe("numerobreve");
        expect(normalizzaNomeColonna("  NUMERO-breve ")).toBe("numerobreve");
    });

    test("ignora gli accenti e i valori assenti", () => {
        expect(normalizzaNomeColonna("Qualità")).toBe("qualita");
        expect(normalizzaNomeColonna(undefined)).toBe("");
        expect(normalizzaNomeColonna(null)).toBe("");
    });
});

describe("risolviColonne", () => {
    const campi: ICampoLista[] = [
        { InternalName: "Title", Title: "Nome" },          // colonna predefinita rinominata
        { InternalName: "Numero", Title: "Numero" },
        { InternalName: "Numero_x0020_breve", Title: "Numero breve" },
        { InternalName: "Created", Title: "Created" },
    ];

    const richieste: IRichiestaColonna[] = [
        { chiave: "nome", candidati: ["Nome", "Contatto"], obbligatoria: true },
        { chiave: "numero", candidati: ["Numero", "NumeroInterno"], obbligatoria: false },
        { chiave: "breve", candidati: ["Numero breve"], obbligatoria: false },
    ];

    test("trova le colonne per nome visualizzato e per nome interno", () => {
        const r = risolviColonne(campi, richieste);
        expect(r.colonne.nome).toBe("Title");
        expect(r.colonne.numero).toBe("Numero");
        expect(r.colonne.breve).toBe("Numero_x0020_breve");
        expect(r.mancanti).toEqual([]);
    });

    test("'Numero' non viene confuso con 'Numero breve' né con 'NumeroInterno'", () => {
        const solo: ICampoLista[] = [
            { InternalName: "NumeroInterno", Title: "NumeroInterno" },
            { InternalName: "Numero_x0020_breve", Title: "Numero breve" },
        ];
        const r = risolviColonne(solo, [{ chiave: "numero", candidati: ["Numero"], obbligatoria: false }]);
        expect(r.colonne.numero).toBeUndefined();
    });

    test("usa i candidati in ordine di preferenza (vecchi nomi come ripiego)", () => {
        const vecchia: ICampoLista[] = [
            { InternalName: "Contatto", Title: "Contatto" },
            { InternalName: "NumeroInterno", Title: "NumeroInterno" },
        ];
        const r = risolviColonne(vecchia, richieste);
        expect(r.colonne.nome).toBe("Contatto");
        expect(r.colonne.numero).toBe("NumeroInterno");
        expect(r.colonne.breve).toBeUndefined();
        expect(r.mancanti).toEqual([]); // "breve" non è obbligatoria
    });

    test("elenca le colonne obbligatorie mancanti con il primo nome candidato", () => {
        const r = risolviColonne([{ InternalName: "Altro", Title: "Altro" }], richieste);
        expect(r.mancanti).toEqual(["Nome"]);
    });

    test("elenco campi vuoto: tutto mancante, nessun errore", () => {
        const r = risolviColonne([], richieste);
        expect(r.mancanti).toEqual(["Nome"]);
        expect(r.colonne.nome).toBeUndefined();
    });

    test("il nome interno ha la precedenza sul nome visualizzato", () => {
        const doppie: ICampoLista[] = [
            { InternalName: "Title", Title: "Nome" },
            { InternalName: "Nome", Title: "Altro" },
        ];
        const r = risolviColonne(doppie, [{ chiave: "nome", candidati: ["Nome"], obbligatoria: true }]);
        expect(r.colonne.nome).toBe("Nome");
    });
});

describe("testoCampo", () => {
    test("stringhe e numeri", () => {
        expect(testoCampo("  101 ")).toBe("101");
        expect(testoCampo(101)).toBe("101");
        expect(testoCampo(0)).toBe("0");
    });

    test("vuoti", () => {
        expect(testoCampo(null)).toBe("");
        expect(testoCampo(undefined)).toBe("");
        expect(testoCampo("   ")).toBe("");
    });

    test("scelta multipla (array) e oggetti non supportati", () => {
        expect(testoCampo(["A", " B ", ""])).toBe("A, B");
        expect(testoCampo({ Title: "x" })).toBe("");
    });
});
