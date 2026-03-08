import * as React from "react";
import { IRubricaUser } from "../types";
import styles from "./Rubrica.module.scss";

interface IPersonCardProps {
    entry: IRubricaUser;
}

const PersonCard: React.FC<IPersonCardProps> = ({ entry }) => {
    const { user, numeroInterno } = entry;

    // Caso servizio: user è null, mostriamo solo mail e numero
    if (!user) {
        return (
            <div className={styles.card}>
                <div className={styles.nome}>Servizio</div>
                {numeroInterno && <div className={styles.numero}>📞 {numeroInterno}</div>}
            </div>
        );
    }

    // Caso utente Graph
    return (
        <div className={styles.card}>
            <div className={styles.nome}>{user.displayName ?? "—"}</div>
            {user.jobTitle && <div className={styles.titolo}>{user.jobTitle}</div>}
            {user.department && <div className={styles.reparto}>{user.department}</div>}
            {user.mail && <div className={styles.mail}>{user.mail}</div>}
            {numeroInterno && <div className={styles.numero}>📞 {numeroInterno}</div>}
        </div>
    );
};

export default PersonCard;