import { getSP } from "../pnpjsConfig";
import { IContattoInterno } from "../types";

const COL_CONTATTO = "Contatto";// NON CAMBIARE DOPO LA CREAZIONE
const COL_INTERNO = "NumeroInterno"; // NON CAMBIARE DOPO LA CREAZIONE

// Il nome (o il percorso) delle liste ora arriva dalle impostazioni del web part.
// Le colonne restano fisse e sono le stesse per tutte le liste.
//   - "Rubrica"                        -> cercata per titolo nel sito corrente
//   - "/sites/intranet/Lists/Esterni"  -> cercata per percorso server-relative
function getLista(riferimento: string) {
    const sp = getSP();
    return riferimento.charAt(0) === "/"
        ? sp.web.getList(riferimento)
        : sp.web.lists.getByTitle(riferimento);
}

// Se la lista non esiste o non è accessibile la chiamata lancia un errore:
// è il chiamante (Rubrica.tsx) a mostrare l'avviso all'utente.
export async function getVociLista(lista: string): Promise<IContattoInterno[]> {
    const items = await getLista(lista)
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
