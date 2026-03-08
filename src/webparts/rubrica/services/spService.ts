import { getSP } from "../pnpjsConfig";
import { IVociLista } from "../types";

const LIST_NAME = "Rubrica";
const COL_CONTATTO = "Contatto";
const COL_INTERNO = "Interno";

export async function getVociLista(): Promise<IVociLista[]> {
    const sp = getSP();

    try {
        const items = await sp.web.lists
            .getByTitle(LIST_NAME)
            .items
            .select(COL_CONTATTO, COL_INTERNO)();

        const voci: IVociLista[] = [];

        for (const item of items) {
            const voce: IVociLista = {
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