type ProfileMeta = Record<string, unknown> | undefined;

function getNow() {
    if (typeof performance !== 'undefined' && typeof performance.now === 'function') {
        return performance.now();
    }
    return Date.now();
}

function roundMs(value: number) {
    return Math.round(value * 100) / 100;
}

export class ProfileLogger {
    private readonly enabled: boolean;
    private readonly scope: string;
    private readonly startedAt: number;
    private lastAt: number;

    constructor(enabled: boolean | undefined, scope: string) {
        this.enabled = Boolean(enabled);
        this.scope = scope;
        this.startedAt = getNow();
        this.lastAt = this.startedAt;
    }

    mark(label: string, meta?: ProfileMeta) {
        if (!this.enabled) {
            return;
        }

        const now = getNow();
        const payload = {
            stepMs: roundMs(now - this.lastAt),
            totalMs: roundMs(now - this.startedAt),
            ...(meta || {}),
        };
        console.log(`[LuckyExcel][profile][${this.scope}] ${label}`, payload);
        this.lastAt = now;
    }

    end(meta?: ProfileMeta) {
        this.mark('done', meta);
    }
}

export function createProfileLogger(enabled: boolean | undefined, scope: string) {
    return new ProfileLogger(enabled, scope);
}
