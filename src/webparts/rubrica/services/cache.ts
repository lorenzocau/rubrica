// Cache in sessionStorage (vale solo per la scheda del browser, sparisce alla chiusura).
// Serve a non richiamare Graph a ogni visita della Homepage.
// Ogni accesso è in try/catch: lo storage può essere bloccato o pieno, e in quel
// caso l'app funziona semplicemente senza cache.

const PREFISSO = "rubrica:";

interface IVoceCache<T> {
    quando: number;
    dati: T;
}

export function leggiCache<T>(chiave: string, ttlMs: number): T | null {
    try {
        const grezzo = window.sessionStorage.getItem(PREFISSO + chiave);
        if (!grezzo) return null;

        const voce = JSON.parse(grezzo) as IVoceCache<T>;
        if (typeof voce.quando !== "number" || Date.now() - voce.quando > ttlMs) return null;

        return voce.dati;
    } catch {
        return null;
    }
}

export function scriviCache<T>(chiave: string, dati: T): void {
    try {
        const voce: IVoceCache<T> = { quando: Date.now(), dati };
        window.sessionStorage.setItem(PREFISSO + chiave, JSON.stringify(voce));
    } catch {
        // storage pieno o non disponibile: si ignora
    }
}
