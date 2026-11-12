declare const API_BASE_URL: string;

export const  CONFIG = {
    API_BASE_URL,
    MAX_FILE_SIZE: 100 * 1024 * 1024,
    MIN_FILE_SIZE: 1024,
    REQUEST_TIMEOUT_MS: 120_000,
    TITLE_MIN: 2,
    TITLE_MAX: 150,
    DESCRIPTION_MAX: 1000,
    TAGS_MAX: 10,
    TAG_MAX_LENGTH: 30,
} as const;
