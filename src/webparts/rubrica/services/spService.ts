import { getSP } from "../pnpjsConfig";
import { IContattoEsterno, IContattoInterno } from "../types";
import { ICampoLista, IRichiestaColonna, risolviColonne, testoCampo } from "../colonne";

// Nomi delle colonne. Si cercano sia per nome interno sia per nome visualizzato
// (vedi colonne.ts), quindi "Numero breve" funziona anche se internamente è "Numero_x0020_breve".
//
// Lista servizi / interni ("Rubrica")
const COL_CONTATTO = "Contatto";// NON CAMBIARE DOPO LA CREAZIONE
const COL_INTERNO = "NumeroInterno"; // NON CAMBIARE DOPO LA CREAZIONE
const COL_REPARTO = "Reparto";

// Lista contatti esterni ("Esterni").
// Per "Nome" e "Numero" si accettano anche i vecchi nomi, così una lista Esterni
// non ancora aggiornata continua a funzionare (senza numero breve).
const COL_ESTERNO_NOME = ["Nome", "Contatto"];
const COL_ESTERNO_NUMERO = ["Numero", "NumeroInterno"];
const COL_ESTERNO_BREVE = ["Numero breve"];

// Il nome (o il percorso) delle liste ora arriva dalle impostazioni del web part.
//   - "Rubrica"                        -> cercata per titolo nel sito corrente
//   - "/sites/intranet/Lists/Esterni"  -> cercata per percorso server-relative
function getLista(riferimento: string) {
    const sp = getSP();
    return riferimento.charAt(0) === "/"
        ? sp.web.getList(riferimento)
        : sp.web.lists.getByTitle(riferimento);
}

interface IRigheLista {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    righe: any[];
    colonne: { [chiave: string]: string | undefined };
}

// Legge le righe di una lista: prima l'elenco delle colonne (per trovare i nomi interni),
// poi gli elementi con solo le colonne che servono.
// Se manca una colonna obbligatoria lancia un errore con .status = 400 e .colonneMancanti,
// che errori.ts trasforma in un messaggio chiaro.
async function leggiRighe(riferimento: string, richieste: IRichiestaColonna[]): Promise<IRigheLista> {
    const lista = getLista(riferimento);

    const campi = (await lista.fields.select("InternalName", "Title")()) as unknown as ICampoLista[];
    const { colonne, mancanti } = risolviColonne(campi, richieste);

    if (mancanti.length > 0) {
        throw Object.assign(
            new Error(`The field or property ${mancanti.map(m => `'${m}'`).join(", ")} does not exist`),
            { status: 400, colonneMancanti: mancanti }
        );
    }

    const daLeggere = richieste
        .map(r => colonne[r.chiave])
        .filter((c): c is string => c !== undefined);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const righe: any[] = await lista.items.select(...daLeggere).top(999)();
    return { righe, colonne };
}

// Se la lista non esiste o non è accessibile la chiamata lancia un errore:
// è il chiamante (Rubrica.tsx) a mostrare l'avviso all'utente.
export async function getVociLista(lista: string): Promise<IContattoInterno[]> {
    const { righe, colonne } = await leggiRighe(lista, [
        { chiave: "contatto", candidati: [COL_CONTATTO], obbligatoria: true },
        { chiave: "interno", candidati: [COL_INTERNO], obbligatoria: true },
        { chiave: "reparto", candidati: [COL_REPARTO], obbligatoria: false },
    ]);

    // la colonna Reparto è facoltativa: senza, i servizi restano visibili ma senza reparto
    if (!colonne.reparto) {
        console.warn(`[Rubrica] La lista "${lista}" non ha la colonna "${COL_REPARTO}": i reparti dei servizi non saranno mostrati.`);
    }

    const voci: IContattoInterno[] = [];

    for (const riga of righe) {
        // testoCampo perché NumeroInterno può arrivare come numero (colonna Numero)
        // e perché le colonne possono essere vuote (null)
        const contatto = testoCampo(riga[colonne.contatto as string]);
        const interno = testoCampo(riga[colonne.interno as string]);
        const reparto = colonne.reparto ? testoCampo(riga[colonne.reparto]) : "";

        if (contatto === "" || interno === "") continue; // riga incompleta

        voci.push({ contatto, interno, reparto });
    }

    return voci;
}

// Lista Esterni: Nome + Numero (intero) + Numero breve.
export async function getVociEsterni(lista: string): Promise<IContattoEsterno[]> {
    const { righe, colonne } = await leggiRighe(lista, [
        { chiave: "nome", candidati: COL_ESTERNO_NOME, obbligatoria: true },
        { chiave: "numero", candidati: COL_ESTERNO_NUMERO, obbligatoria: false },
        { chiave: "numeroBreve", candidati: COL_ESTERNO_BREVE, obbligatoria: false },
    ]);

    // serve almeno una delle due colonne con i numeri
    if (!colonne.numero && !colonne.numeroBreve) {
        throw Object.assign(
            new Error("The field or property 'Numero' does not exist"),
            { status: 400, colonneMancanti: [COL_ESTERNO_NUMERO[0], COL_ESTERNO_BREVE[0]] }
        );
    }

    if (!colonne.numeroBreve) {
        console.warn(`[Rubrica] La lista "${lista}" non ha la colonna "${COL_ESTERNO_BREVE[0]}": i numeri brevi degli esterni non saranno disponibili.`);
    }

    const voci: IContattoEsterno[] = [];

    for (const riga of righe) {
        const nome = testoCampo(riga[colonne.nome as string]);
        const numero = colonne.numero ? testoCampo(riga[colonne.numero]) : "";
        const numeroBreve = colonne.numeroBreve ? testoCampo(riga[colonne.numeroBreve]) : "";

        if (nome === "" || (numero === "" && numeroBreve === "")) continue; // riga incompleta

        voci.push({ nome, numero, numeroBreve });
    }

    return voci;
}
