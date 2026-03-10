import { getSP } from "../pnpjsConfig";
import { IContattoInterno } from "../types";

const LIST_NAME = "Rubrica";// NON CAMBIARE DOPO LA CREAZION
const COL_CONTATTO = "Contatto";// NON CAMBIARE DOPO LA CREAZION
const COL_INTERNO = "NumeroInterno"; // NON CAMBIARE DOPO LA CREAZIONE

export async function getVociLista(): Promise<IContattoInterno[]> {
    const sp = getSP();

    try {
        const items = await sp.web.lists
            .getByTitle(LIST_NAME)
            .items
            .select(COL_CONTATTO, COL_INTERNO)
            .top(999)();

        const voci: IContattoInterno[] = [];

        for (const item of items) {
            const voce: IContattoInterno = {
                contatto: item[COL_CONTATTO],
                interno: item[COL_INTERNO],
            };
            voci.push(voce);
        }
        
        return voci;

    } catch (e) {
        console.warn(`Lista "${LIST_NAME}" non trovata o non accessibile.`, e);
        return [];
    }
}