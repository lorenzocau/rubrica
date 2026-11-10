import * as React from "react";
import { Persona, PersonaPresence, PersonaSize } from "@fluentui/react/lib/Persona";
import { Link } from "@fluentui/react/lib/Link";
import { Icon } from "@fluentui/react/lib/Icon";
import { IconButton } from "@fluentui/react/lib/Button";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import { IOpzioniCard, IRubricaEntry } from "../types";
import { descrizioneUtente, dividiNumeri, SEPARATORE_INTERNI } from "../rubricaLogic";

interface IPersonCardProps {
    entry: IRubricaEntry;
    opzioni: IOpzioniCard;
}

// --- stili ---
const cardStyle = mergeStyles({
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "8px 16px",
    borderBottom: "1px solid #edebe9",
    ":hover": { backgroundColor: "#f3f2f1" }
});

const destraStyle = mergeStyles({
    display: "flex",
    alignItems: "center",
    flexShrink: 0
});

// Tre slot fissi, sempre nello stesso ordine: interno | numero breve | cellulare intero.
// Se un dato manca lo slot resta vuoto, così le colonne sono allineate in tutta la lista.
// Cifre di larghezza uguale (tabular-nums) per confrontare i numeri tra righe diverse.
const numeriStyle = mergeStyles({
    display: "grid",
    gridTemplateColumns: "72px 72px 128px",
    columnGap: "4px",
    alignItems: "center",
    fontVariantNumeric: "tabular-nums"
});

// slot di un numero breve (interno o breve del cellulare): icona + cifre in evidenza
const slotBreveStyle = mergeStyles({
    display: "flex",
    alignItems: "center",
    fontWeight: 600,
    fontSize: 14,
    color: "#000000",
    whiteSpace: "nowrap"
});

const iconaStyle = mergeStyles({
    fontSize: 12,
    color: "#605e5c",
    marginRight: 5,
    flexShrink: 0
});

// slot del cellulare intero: secondario, in grigio
const slotCellulareStyle = mergeStyles({
    fontSize: 12,
    color: "#605e5c",
    whiteSpace: "nowrap"
});

// Spazio fisso del pulsante chat: c'è sempre, anche dove il pulsante manca (servizi, esterni,
// utenti senza mail), così i numeri finiscono sullo stesso bordo destro in tutta la lista.
const chatSlotStyle = mergeStyles({
    width: 32,
    height: 32,
    marginLeft: 4,
    flexShrink: 0
});

// Link dei numeri in grigio neutro (il rosso del tema su 12px è poco leggibile in una lista
// fitta); la sottolineatura compare al passaggio del mouse.
const linkNumeroStyles = {
    root: {
        color: "#605e5c",
        ":hover": { color: "#323130", textDecoration: "underline" }
    }
};

// "333 1234567" -> "tel:3331234567"
function hrefTel(numero: string): string {
    return `tel:${numero.replace(/\s+/g, "")}`;
}

// Un numero con il link per chiamare, oppure solo il testo.
function renderNumero(numero: string, cliccabile: boolean): React.ReactNode {
    return cliccabile ? <Link href={hrefTel(numero)} styles={linkNumeroStyles}>{numero}</Link> : numero;
}

// Più numeri (uniti con " / " in rubricaLogic): ognuno ha il suo link.
function renderNumeri(numeri: string, cliccabile: boolean): React.ReactNode {
    return dividiNumeri(numeri).map((n, i) => (
        <React.Fragment key={i}>
            {i > 0 && SEPARATORE_INTERNI}
            {renderNumero(n, cliccabile)}
        </React.Fragment>
    ));
}

interface IBloccoNumeriProps {
    interno?: string;           // icona cornetta
    breve?: string;             // icona telefono: numero breve del cellulare
    cellulare?: React.ReactNode; // cellulare intero, in grigio
}

// I tre slot dei numeri. Un valore vuoto o assente lascia lo slot vuoto.
function BloccoNumeri({ interno, breve, cellulare }: IBloccoNumeriProps) {
    return (
        <div className={numeriStyle}>
            <div className={slotBreveStyle} title={interno ? "Interno" : undefined}>
                {interno && <><Icon iconName="Phone" className={iconaStyle} aria-hidden="true" />{interno}</>}
            </div>
            <div className={slotBreveStyle} title={breve ? "Numero breve del cellulare" : undefined}>
                {breve && <><Icon iconName="CellPhone" className={iconaStyle} aria-hidden="true" />{breve}</>}
            </div>
            <div className={slotCellulareStyle} title={cellulare ? "Cellulare" : undefined}>
                {cellulare}
            </div>
        </div>
    );
}

function PersonCard({ entry, opzioni }: IPersonCardProps) {

    if (entry.tipo === "utente") {
        const user = entry.user;

        // "Reparto - Job title" di AD (department e jobTitle)
        const secondaryText = opzioni.mostraReparto ? descrizioneUtente(user) : "";

        // la mail non si mostra: serve solo per il link alla chat di Teams
        const mail = user.mail ?? undefined;

        // cellulare aziendale (mobilePhone) e suo numero breve (faxNumber)
        const cellulare = opzioni.mostraCellulare ? (user.mobilePhone ?? "").trim() : "";
        const cellulareBreve = opzioni.mostraCellulare ? (user.faxNumber ?? "").trim() : "";

        return (
            <div className={cardStyle}>
                <Persona
                    text={user.displayName ?? "—"}
                    secondaryText={secondaryText || undefined}
                    size={PersonaSize.size40}
                    presence={opzioni.mostraPresenza ? PersonaPresence.online : PersonaPresence.none}
                />
                <div className={destraStyle}>
                    <BloccoNumeri
                        interno={entry.interno ?? undefined}
                        breve={cellulareBreve}
                        cellulare={cellulare !== "" ? renderNumero(cellulare, opzioni.linkCliccabili) : undefined}
                    />
                    <div className={chatSlotStyle}>
                        {opzioni.linkCliccabili && mail && (
                            <IconButton
                                iconProps={{ iconName: "Chat" }}
                                title="Apri una chat in Teams"
                                ariaLabel={`Apri una chat in Teams con ${user.displayName ?? mail}`}
                                href={`https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(mail)}`}
                                target="_blank"
                                rel="noopener noreferrer"
                            />
                        )}
                    </div>
                </div>
            </div>
        );
    }

    // servizio: nessun account AD, nome, reparto (al posto del job title) e interno
    if (entry.tipo === "servizio") {
        const secondaryText = opzioni.mostraReparto && entry.reparto !== "" ? entry.reparto : "Servizio";

        return (
            <div className={cardStyle}>
                <Persona
                    text={entry.nome}
                    secondaryText={secondaryText}
                    size={PersonaSize.size40}
                />
                <div className={destraStyle}>
                    <BloccoNumeri interno={entry.interno} />
                    <div className={chatSlotStyle} />
                </div>
            </div>
        );
    }

    // esterno: numero breve nello slot del breve, numero intero in quello del cellulare
    return (
        <div className={cardStyle}>
            <Persona
                text={entry.nome}
                secondaryText="Esterno"
                size={PersonaSize.size40}
            />
            <div className={destraStyle}>
                <BloccoNumeri
                    breve={entry.numeroBreve}
                    cellulare={entry.numero !== "" ? renderNumeri(entry.numero, opzioni.linkCliccabili) : undefined}
                />
                <div className={chatSlotStyle} />
            </div>
        </div>
    );
}

export default PersonCard;