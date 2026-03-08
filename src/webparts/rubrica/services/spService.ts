import { getSP } from "../pnpjsConfig";
import { IVociLista } from "../types";

// Nome della lista e delle colonne SharePoint.
// Centralizzati qui per evitare "magic strings" sparse nel codice —
// se un giorno cambia il nome basta modificarlo in un posto solo.
const LIST_NAME = "Rubrica";
const COL_MAIL = "Mail";
const COL_NUMERO_INTERNO = "NumeroInterno";

// Legge tutti gli item della lista e li restituisce come IVociLista[].
// Non gestisce errori: se la chiamata fallisce l'eccezione sale a Rubrica.tsx
// che sa come mostrarla all'utente.
export const getVociLista = async (): Promise<IVociLista[]> => {
    const sp = getSP();

    // .select() → chiede a SharePoint solo i campi che ci servono,
    //             evitando di ricevere decine di colonne inutili
    // ()        → esegue la query (sintassi PnPjs v3)
    const items = await sp.web.lists
        .getByTitle(LIST_NAME)
        .items
        .select(COL_MAIL, COL_NUMERO_INTERNO)();

    // Mappiamo la risposta grezza di SharePoint nel nostro tipo IVociLista.
    // SharePoint restituisce i nomi delle colonne così come sono,
    // quindi "Mail" e "NumeroInterno" corrispondono direttamente.
    return items.map((item) => ({
        mail: item[COL_MAIL],
        numeroInterno: item[COL_NUMERO_INTERNO],
    }));
};