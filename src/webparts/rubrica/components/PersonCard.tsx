import * as React from "react";
import { Persona, PersonaPresence, PersonaSize } from "@fluentui/react/lib/Persona";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import { IRubricaEntry } from "../types";

interface IPersonCardProps {
    entry: IRubricaEntry;
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

const numeroStyle = mergeStyles({
    fontWeight: 600,
    fontSize: 16,
    color: "#aa1414",
    minWidth: 80,
    textAlign: "right",
    flexShrink: 0
});

function PersonCard({ entry }: IPersonCardProps) {

    if (entry.tipo === "utente") {
        return (
            <div className={cardStyle}>
                <Persona
                    text={entry.user.displayName ?? "—"}
                    secondaryText={`${entry.user.department} - ${entry.user.jobTitle}`}
                    tertiaryText={entry.user.mail ?? undefined} // size56 o inferiore
                    size={PersonaSize.size40}
                    presence={PersonaPresence.online}

                />
                <div className={numeroStyle}>
                    {entry.interno ? `int. ${entry.interno}` : "—"}
                </div>
            </div>
        );
    }
    // PersonaSize.size100 = fino a tertiary
    // PersonaSize.size72 = fino a secondary

    // servizio: nessun account AD, solo nome e numero
    return (
        <div className={cardStyle}>
            <Persona
                text={entry.nome}
                secondaryText="Servizio"
                tertiaryText="Miao miao"
                size={PersonaSize.size48}

            />
            <div className={numeroStyle}>
                {`int. ${entry.interno}`}
            </div>
        </div>
    );
}

export default PersonCard;