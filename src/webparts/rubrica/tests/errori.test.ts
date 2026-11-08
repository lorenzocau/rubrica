import { leggiErrore, erroreGraph, erroreLista, erroreImprevisto, AVVISO_GRAPH_VUOTO } from "../errori";

// Imita HttpRequestError di PnPjs: un Error con .status
function httpError(status: number, message = "Error making HTTP request"): Error {
    return Object.assign(new Error(message), { status });
}

describe("leggiErrore", () => {
    test("legge il codice da .status", () => {
        expect(leggiErrore(httpError(403)).stato).toBe(403);
    });

    test("legge il codice da response.status", () => {
        expect(leggiErrore({ message: "x", response: { status: 404 } }).stato).toBe(404);
    });

    test("ripiego: cerca il codice nel testo del messaggio", () => {
        expect(leggiErrore(new Error("Error making HTTP request: 404")).stato).toBe(404);
    });

    test("un numero qualsiasi nel testo non è un codice HTTP", () => {
        expect(leggiErrore(new Error("trovati 404 elementi")).stato).toBeUndefined();
        expect(leggiErrore(new Error("id 12345")).stato).toBeUndefined();
    });

    test("TypeError (fetch senza rete) è un errore di rete", () => {
        const d = leggiErrore(new TypeError("Failed to fetch"));
        expect(d.rete).toBe(true);
        expect(d.stato).toBeUndefined();
    });

    test("un errore con codice HTTP non è di rete", () => {
        expect(leggiErrore(httpError(500)).rete).toBe(false);
    });

    test("valori strani non lanciano eccezioni", () => {
        expect(leggiErrore(undefined).testo).toBe("");
        expect(leggiErrore(null).rete).toBe(false);
        expect(leggiErrore("boom").stato).toBeUndefined();
        expect(leggiErrore(42).testo).toBe("");
    });

    test("il testo troppo lungo viene accorciato", () => {
        expect(leggiErrore(new Error("a".repeat(500))).testo.length).toBeLessThan(200);
    });
});

describe("erroreGraph", () => {
    test("403: indica il permesso User.Read.All e il codice", () => {
        const m = erroreGraph(httpError(403));
        expect(m).toContain("User.Read.All");
        expect(m).toContain("403");
    });

    test("401: suggerisce di ricaricare", () => {
        expect(erroreGraph(httpError(401))).toContain("ricarica");
    });

    test("429 e 5xx hanno messaggi propri", () => {
        expect(erroreGraph(httpError(429))).toContain("limitando");
        expect(erroreGraph(httpError(503))).toContain("non sta rispondendo");
    });

    test("rete assente", () => {
        expect(erroreGraph(new TypeError("Failed to fetch"))).toContain("raggiungere");
    });

    test("causa sconosciuta: include il dettaglio tecnico", () => {
        expect(erroreGraph(new Error("qualcosa di strano"))).toContain("qualcosa di strano");
    });

    test("dice sempre cosa resta in rubrica", () => {
        expect(erroreGraph(httpError(403))).toContain("servizi ed esterni");
        expect(erroreGraph(new Error("x"))).toContain("servizi ed esterni");
    });

    test("l'avviso per Graph vuoto cita il permesso", () => {
        expect(AVVISO_GRAPH_VUOTO).toContain("User.Read.All");
    });
});

describe("erroreLista", () => {
    const conseguenza = "I servizi non sono disponibili.";

    test("404: lista non trovata, con nome e suggerimento sul percorso", () => {
        const m = erroreLista(httpError(404), "Rubrica", conseguenza);
        expect(m).toContain('"Rubrica"');
        expect(m).toContain("non trovata");
        expect(m).toContain("percorso");
        expect(m).toContain(conseguenza);
    });

    test("403: mancano i permessi di lettura", () => {
        expect(erroreLista(httpError(403), "Esterni", conseguenza)).toContain("permessi");
        expect(erroreLista(httpError(401), "Esterni", conseguenza)).toContain("permessi");
    });

    test("400 o 'does not exist' sul campo: colonne mancanti", () => {
        expect(erroreLista(httpError(400), "Rubrica", conseguenza)).toContain("NumeroInterno");
        const e = new Error("The field or property 'Contatto' does not exist");
        expect(erroreLista(e, "Rubrica", conseguenza)).toContain("colonne");
    });

    test("404 con 'does not exist' resta 'lista non trovata'", () => {
        const e = httpError(404, "The list does not exist");
        expect(erroreLista(e, "Rubrica", conseguenza)).toContain("non trovata");
    });

    test("rete assente e 5xx", () => {
        expect(erroreLista(new TypeError("Failed to fetch"), "Rubrica", conseguenza)).toContain("raggiungere");
        expect(erroreLista(httpError(500), "Rubrica", conseguenza)).toContain("non sta rispondendo");
    });

    test("causa sconosciuta: dettaglio tecnico e conseguenza", () => {
        const m = erroreLista(new Error("boh"), "Rubrica", conseguenza);
        expect(m).toContain("boh");
        expect(m).toContain(conseguenza);
    });
});

describe("erroreImprevisto", () => {
    test("include il messaggio dell'errore", () => {
        expect(erroreImprevisto(new Error("cannot read x"))).toContain("cannot read x");
    });

    test("funziona anche senza messaggio", () => {
        expect(erroreImprevisto(undefined)).toContain("Errore imprevisto");
    });
});