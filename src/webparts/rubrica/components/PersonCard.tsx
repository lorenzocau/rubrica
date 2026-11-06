import * as React from "react";
import { Persona, PersonaPresence, PersonaSize } from "@fluentui/react/lib/Persona";
import { Link } from "@fluentui/react/lib/Link";
import { IconButton } from "@fluentui/react/lib/Button";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import { IOpzioniCard, IRubricaEntry } from "../types";

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

function PersonCard({ entry, opzioni }: IPersonCardProps) {

    if (entry.tipo === "utente") {
        const user = entry.user;

        // "Office" di AD: contiene già reparto e ruolo insieme, si mostra così com'è
        const secondaryText = opzioni.mostraReparto ? (user.officeLocation ?? "").trim() : "";

        const mail = user.mail ?? undefined;
        const cellulare = opzioni.mostraCellulare ? (user.mobilePhone ?? undefined) : undefined;

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
                            {entry.interno ? `int. ${entry.interno}` : "—"}
                        </div>
                        {cellulare && (
                            <div className={cellulareStyle}>
                                {opzioni.linkCliccabili
                                    ? <Link href={`tel:${cellulare.replace(/\s+/g, "")}`}>{cellulare}</Link>
                                    : cellulare}
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

    // servizio / esterno: nessun account AD, solo nome e numero
    return (
        <div className={cardStyle}>
            <Persona
                text={entry.nome}
                secondaryText={entry.tipo === "esterno" ? "Esterno" : "Servizio"}
                size={PersonaSize.size48}
            />
            <div className={numeriStyle}>
                <div className={numeroStyle}>
                    {`int. ${entry.interno}`}
                </div>
            </div>
        </div>
    );
}

export default PersonCard;
