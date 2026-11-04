import { getSP } from "../pnpjsConfig";
import { IContattoInterno } from "../types";

export const LIST_NAME = "Rubrica"; // NON CAMBIARE DOPO LA CREAZIONE
const COL_CONTATTO = "Contatto";// NON CAMBIARE DOPO LA CREAZIONE
const COL_INTERNO = "NumeroInterno"; // NON CAMBIARE DOPO LA CREAZIONE

// Se la lista non esiste o non è accessibile la chiamata lancia un errore:
// è il chiamante (Rubrica.tsx) a mostrare l'avviso all'utente.
export async function getVociLista(): Promise<IContattoInterno[]> {
    const sp = getSP();

    const items = await sp.web.lists
        .getByTitle(LIST_NAME)
        .items
        .select(COL_CONTATTO, COL_INTERNO)
        .top(999)();

    const voci: IContattoInterno[] = [];

    for (const item of items) {
        // String() perché NumeroInterno può arrivare come numero (colonna Numero)
        // e perché entrambe le colonne possono essere vuote (null)
        const contatto = String(item[COL_CONTATTO] ?? "").trim();
        const interno = String(item[COL_INTERNO] ?? "").trim();

        if (contatto === "" || interno === "") continue; // riga incompleta

        voci.push({ contatto, interno });
    }

    return voci;
}