import { IRubricaWebPartProps } from "./types";

// Valori di default. DEVONO coincidere con "properties" in RubricaWebPart.manifest.json
// (il manifest vale per le NUOVE istanze del web part, questi per quelle già in pagina).
export const PROPRIETA_DEFAULT: IRubricaWebPartProps = {
  titolo: "",
  placeholderRicerca: "Cerca per nome, reparto, interno...",
  risultatiPerPagina: 30,
  mostraElencoCompleto: false,

  listaServizi: "Rubrica",
  listaEsterni: "Esterni",

  mostraEsterni: true,
  esterniInizialmenteVisibili: true,

  soloMembri: true,
  nascondiSenzaMail: true,

  mostraContatore: true,
  mostraPresenza: true,
  mostraMail: true,
  mostraReparto: true,
  mostraCellulare: true,
  linkCliccabili: true,

  usaDatiMock: false,
};

// Un web part già messo in pagina prima di questa versione ha proprietà salvate
// vecchie (o nessuna): per ogni chiave mancante o di tipo sbagliato si usa il default.
export function risolviProprieta(salvate?: Partial<IRubricaWebPartProps> | null): IRubricaWebPartProps {
  const risolte: IRubricaWebPartProps = { ...PROPRIETA_DEFAULT };
  if (!salvate) return risolte;

  const dest = risolte as unknown as Record<string, unknown>;
  const src = salvate as unknown as Record<string, unknown>;
  const def = PROPRIETA_DEFAULT as unknown as Record<string, unknown>;

  Object.keys(def).forEach(chiave => {
    const valore = src[chiave];
    if (valore !== undefined && valore !== null && typeof valore === typeof def[chiave]) {
      dest[chiave] = valore;
    }
  });

  return risolte;
}
