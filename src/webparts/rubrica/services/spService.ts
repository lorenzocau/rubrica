import { getSP } from "../pnpjsConfig";
import { IVociLista } from "../types";

// Nome della lista e delle colonne SharePoint.
// Centralizzati qui per evitare "magic strings" sparse nel codice —
// se un giorno cambia il nome basta modificarlo in un posto solo.
const LIST_NAME = "Rubrica";
const COL_MAIL = "Mail";
const COL_NUMERO_INTERNO = "NumeroInterno";

// Legge tutti gli item della lista e li restituisce come IVociLista[].
// Se la lista non esiste o la chiamata fallisce, restituisce un array vuoto
// — la rubrica mostrerà comunque tutti gli utenti Graph senza numeri interni.
export const getVociLista = async (): Promise<IVociLista[]> => {
    const sp = getSP();

    try {
        const items = await sp.web.lists
            .getByTitle(LIST_NAME)
            .items
            .select(COL_MAIL, COL_NUMERO_INTERNO)();

        return items.map((item) => ({
            mail: item[COL_MAIL],
            numeroInterno: item[COL_NUMERO_INTERNO],
        }));
    } catch (e) {
        console.warn(`Lista "${LIST_NAME}" non trovata o non accessibile. I numeri interni non saranno disponibili.`, e);
        return [];
    }
};