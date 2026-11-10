// Trova le colonne di una lista SharePoint a partire dal nome "umano" (es. "Numero breve").
// Funzioni pure, senza dipendenze da SPFx: si possono testare.
//
// Perché serve: il nome INTERNO di una colonna non coincide con quello che si vede.
// "Numero breve" creata da interfaccia diventa "Numero_x0020_breve", e la colonna
// predefinita "Title" può essere stata rinominata "Nome" restando "Title" all'interno.
// Confrontando sia il nome interno sia il nome visualizzato, senza badare a
// maiuscole, spazi, "_x0020_" e accenti, funziona in tutti questi casi.

export interface ICampoLista {
    InternalName?: string;
    Title?: string;
}

export interface IRichiestaColonna {
    chiave: string;        // nome logico usato dal codice (es. "numeroBreve")
    candidati: string[];   // nomi accettati, in ordine di preferenza
    obbligatoria: boolean;
}

export interface IColonneRisolte {
    colonne: { [chiave: string]: string | undefined }; // chiave -> nome interno trovato
    mancanti: string[];                                // primo nome candidato delle obbligatorie non trovate
}

// "Numero_x0020_breve", "Numero breve" e "NumeroBreve" diventano tutti "numerobreve".
export function normalizzaNomeColonna(nome?: string | null): string {
    return (nome ?? "")
        .replace(/_x([0-9a-f]{4})_/gi, (_m, hex: string) => String.fromCharCode(parseInt(hex, 16)))
        .normalize("NFD")
        .replace(/[̀-ͯ]/g, "")
        .replace(/[^a-z0-9]/gi, "")
        .toLowerCase();
}

export function risolviColonne(campi: ICampoLista[], richieste: IRichiestaColonna[]): IColonneRisolte {
    const colonne: { [chiave: string]: string | undefined } = {};
    const mancanti: string[] = [];

    for (const richiesta of richieste) {
        let trovata: string | undefined;

        for (const candidato of richiesta.candidati) {
            const cercato = normalizzaNomeColonna(candidato);
            if (cercato === "") continue;

            // prima il nome interno, poi quello visualizzato
            const campo =
                campi.find(c => normalizzaNomeColonna(c.InternalName) === cercato) ??
                campi.find(c => normalizzaNomeColonna(c.Title) === cercato);

            if (campo && campo.InternalName) {
                trovata = campo.InternalName;
                break;
            }
        }

        colonne[richiesta.chiave] = trovata;
        if (!trovata && richiesta.obbligatoria) mancanti.push(richiesta.candidati[0]);
    }

    return { colonne, mancanti };
}

// Legge il valore di una colonna come testo senza fidarsi del tipo:
// SharePoint può restituire stringhe, numeri (colonna Numero), null, array (scelta multipla)
// oppure oggetti (ricerca, persona) che qui non sono supportati e valgono "".
export function testoCampo(valore: unknown): string {
    if (valore === null || valore === undefined) return "";
    if (typeof valore === "string") return valore.trim();
    if (typeof valore === "number" || typeof valore === "boolean") return String(valore);
    if (Array.isArray(valore)) {
        return valore.map(v => testoCampo(v)).filter(t => t !== "").join(", ");
    }
    return "";
}
