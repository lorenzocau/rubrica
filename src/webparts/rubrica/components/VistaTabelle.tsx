import * as React from "react";
import { useMemo, useState } from "react";
import { Checkbox } from "@fluentui/react/lib/Checkbox";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Text } from "@fluentui/react/lib/Text";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import { IRubricaEntry } from "../types";
import { costruisciTabelle, formattaDataStampa, generaCsv, generaHtmlStampa, ITabella, nomeFileCsv } from "../tabelle";

interface IVistaTabelleProps {
    rubrica: IRubricaEntry[];     // elenco completo, senza filtri (i filtri di questa vista sono in tabelle.ts)
    ricerca: string;              // "" = tutto
    titolo: string;               // titolo del web part, usato nella stampa
    esterniDisponibili: boolean;  // la lista Esterni è attiva e letta
    mostraCellulare: boolean;
    mostraReparto: boolean;
}

// --- stili ---
const barraStyle = mergeStyles({
    display: "flex",
    flexWrap: "wrap",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 12,
    margin: "8px 0 4px"
});

const sezioniStyle = mergeStyles({
    display: "flex",
    flexWrap: "wrap",
    alignItems: "center",
    gap: 16
});

const azioniStyle = mergeStyles({
    display: "flex",
    gap: 8
});

const tabellaStyle = mergeStyles({
    width: "100%",
    borderCollapse: "collapse",
    fontSize: 13,
    selectors: {
        "& th": {
            textAlign: "left",
            fontWeight: 600,
            padding: "6px 10px",
            backgroundColor: "#f3f2f1",
            borderBottom: "2px solid #8a8886"
        },
        "& td": {
            padding: "5px 10px",
            borderBottom: "1px solid #edebe9",
            verticalAlign: "top"
        },
        "& td.num": {
            whiteSpace: "nowrap",
            fontVariantNumeric: "tabular-nums"
        },
        "& tbody tr:hover": {
            backgroundColor: "#f3f2f1"
        }
    }
});

const NUMERICHE = ["interno", "cellulare", "cellulareBreve"];

function TabellaDati({ tabella }: { tabella: ITabella }) {
    return (
        <div style={{ marginBottom: 24 }}>
            <Text as="h3" variant="large" block style={{ margin: "16px 0 6px" }}>
                {tabella.titolo} <span style={{ color: "#605e5c", fontWeight: 400 }}>({tabella.righe.length})</span>
            </Text>

            {tabella.righe.length === 0 ? (
                <Text variant="small" style={{ color: "#605e5c" }}>Nessun contatto.</Text>
            ) : (
                <table className={tabellaStyle} aria-label={tabella.titolo}>
                    <thead>
                        <tr>
                            {tabella.colonne.map(c => <th key={c.chiave}>{c.titolo}</th>)}
                        </tr>
                    </thead>
                    <tbody>
                        {tabella.righe.map((riga, i) => (
                            <tr key={i}>
                                {tabella.colonne.map(c => (
                                    <td key={c.chiave} className={NUMERICHE.indexOf(c.chiave) !== -1 ? "num" : undefined}>
                                        {riga[c.chiave] ?? ""}
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            )}
        </div>
    );
}

// Scarica un testo come file (con il BOM UTF-8, così Excel legge bene gli accenti).
function scaricaCsv(nomeFile: string, contenuto: string): void {
    const blob = new Blob(["﻿" + contenuto], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = nomeFile;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function VistaTabelle({ rubrica, ricerca, titolo, esterniDisponibili, mostraCellulare, mostraReparto }: IVistaTabelleProps) {
    const [conDipendenti, setConDipendenti] = useState<boolean>(true);
    const [conServizi, setConServizi] = useState<boolean>(true);
    const [conEsterni, setConEsterni] = useState<boolean>(true);
    const [errore, setErrore] = useState<string>("");

    const ricercaPulita = ricerca.trim();

    const tabelle = useMemo<ITabella[]>(
        () => costruisciTabelle(rubrica, ricerca, {
            dipendenti: conDipendenti,
            servizi: conServizi,
            esterni: esterniDisponibili && conEsterni,
            mostraCellulare,
            mostraReparto,
        }),
        [rubrica, ricerca, conDipendenti, conServizi, conEsterni, esterniDisponibili, mostraCellulare, mostraReparto]
    );

    const totale = tabelle.reduce((somma, t) => somma + t.righe.length, 0);

    // Stampa in una finestra pulita: solo le tabelle, senza la pagina SharePoint intorno.
    function stampa(): void {
        setErrore("");
        const html = generaHtmlStampa(tabelle, {
            titolo: titolo.trim() !== "" ? titolo.trim() : "Rubrica",
            dataStampa: formattaDataStampa(new Date()),
            filtro: ricercaPulita,
        });

        const finestra = window.open("", "_blank");
        if (!finestra) {
            setErrore("Il browser ha bloccato la finestra di stampa: consenti i popup per questo sito e riprova.");
            return;
        }
        finestra.document.open();
        finestra.document.write(html);
        finestra.document.close();
        // piccola attesa: la finestra deve aver finito di disegnare la pagina
        setTimeout(() => {
            finestra.focus();
            finestra.print();
        }, 300);
    }

    function esporta(): void {
        scaricaCsv(nomeFileCsv(new Date()), generaCsv(tabelle));
    }

    return (
        <>
            <div className={barraStyle}>
                <div className={sezioniStyle}>
                    <Checkbox label="Dipendenti" checked={conDipendenti} onChange={(_e, c) => setConDipendenti(!!c)} />
                    <Checkbox label="Servizi" checked={conServizi} onChange={(_e, c) => setConServizi(!!c)} />
                    {esterniDisponibili && (
                        <Checkbox label="Esterni" checked={conEsterni} onChange={(_e, c) => setConEsterni(!!c)} />
                    )}
                </div>
                <div className={azioniStyle}>
                    <DefaultButton iconProps={{ iconName: "Print" }} text="Stampa" disabled={totale === 0} onClick={stampa} />
                    <DefaultButton iconProps={{ iconName: "Download" }} text="Esporta CSV" disabled={totale === 0} onClick={esporta} />
                </div>
            </div>

            {errore !== "" && (
                <MessageBar messageBarType={MessageBarType.warning} onDismiss={() => setErrore("")}>{errore}</MessageBar>
            )}

            {totale === 0 && (
                <Text variant="medium" style={{ color: "#605e5c", display: "block", margin: "12px 0" }}>
                    {ricercaPulita !== "" ? `Nessun risultato per “${ricercaPulita}”.` : "Nessun contatto disponibile."}
                </Text>
            )}

            {totale > 0 && tabelle.map(t => <TabellaDati key={t.id} tabella={t} />)}
        </>
    );
}

export default VistaTabelle;
