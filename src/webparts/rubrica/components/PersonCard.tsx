import * as React from "react";
import { Persona, PersonaPresence, PersonaSize } from "@fluentui/react/lib/Persona";
import { Link } from "@fluentui/react/lib/Link";
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

const numeriStyle = mergeStyles({
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-end",
    minWidth: 80
});

const numeroStyle = mergeStyles({
    fontWeight: 600,
    fontSize: 12,
    color: "#000000",
    textAlign: "right"
});

const cellulareStyle = mergeStyles({
    fontSize: 11,
    color: "#605e5c",
    marginTop: 2,
    textAlign: "right"
});

// "333 1234567" -> "tel:3331234567"
function hrefTel(numero: string): string {
    return `tel:${numero.replace(/\s+/g, "")}`;
}

// Un numero con il link per chiamare, oppure solo il testo.
function renderNumero(numero: string, cliccabile: boolean): React.ReactNode {
    return cliccabile ? <Link href={hrefTel(numero)}>{numero}</Link> : numero;
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

function PersonCard({ entry, opzioni }: IPersonCardProps) {

    if (entry.tipo === "utente") {
        const user = entry.user;

        // "Reparto - Job title" di AD (department e jobTitle)
        const secondaryText = opzioni.mostraReparto ? descrizioneUtente(user) : "";

        const mail = user.mail ?? undefined;

        // cellulare aziendale (mobilePhone) e suo numero breve (faxNumber)
        const cellulare = opzioni.mostraCellulare ? (user.mobilePhone ?? "").trim() : "";
        const cellulareBreve = opzioni.mostraCellulare ? (user.faxNumber ?? "").trim() : "";

        const mailVisibile = opzioni.mostraMail && !!mail;
        const mailCliccabile = mailVisibile && opzioni.linkCliccabili;

        return (
            <div className={cardStyle}>
                <Persona
                    text={user.displayName ?? "—"}
                    secondaryText={secondaryText || undefined}
                    tertiaryText={mailVisibile ? mail : undefined} // size56 o inferiore
                    onRenderTertiaryText={mailCliccabile
                        ? () => <Link href={`mailto:${mail}`}>{mail}</Link>
                        : undefined}
                    size={PersonaSize.size40}
                    presence={opzioni.mostraPresenza ? PersonaPresence.online : PersonaPresence.none}
                />
                <div className={destraStyle}>
                    <div className={numeriStyle}>
                        <div className={numeroStyle}>
                            {entry.interno ? entry.interno : "—"}
                        </div>
                        {(cellulare !== "" || cellulareBreve !== "") && (
                            <div className={cellulareStyle}>
                                {cellulare !== "" && renderNumero(cellulare, opzioni.linkCliccabili)}
                                {cellulare !== "" && cellulareBreve !== "" && " - "}
                                {cellulareBreve}
                            </div>
                        )}
                    </div>
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
        );
    }
    // PersonaSize.size100 = fino a tertiary
    // PersonaSize.size72 = fino a secondary

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
                <div className={numeriStyle}>
                    <div className={numeroStyle}>
                        {entry.interno}
                    </div>
                </div>
            </div>
        );
    }

    // esterno: sopra il numero breve (con prefisso), sotto il numero per intero.
    // Se manca uno dei due, quello che c'è prende il posto principale.
    const haBreve = entry.numeroBreve !== "";
    const haNumero = entry.numero !== "";

    return (
        <div className={cardStyle}>
            <Persona
                text={entry.nome}
                secondaryText="Esterno"
                size={PersonaSize.size40}
            />
            <div className={numeriStyle}>
                {haBreve && (
                    <div className={numeroStyle}>
                        {`nr. breve ${entry.numeroBreve}`}
                    </div>
                )}
                {haNumero && (
                    <div className={haBreve ? cellulareStyle : numeroStyle}>
                        {renderNumeri(entry.numero, opzioni.linkCliccabili)}
                    </div>
                )}
            </div>
        </div>
    );
}

export default PersonCard;
