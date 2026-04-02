export interface SlideContent {
    title: string;
    bullets: string[];
    images: string[]; // Base64 or URLs of extracted/generated images
    level?: number;
    breadcrumb?: string;
}

export interface DocumentData {
    title: string;
    slides: SlideContent[];
}
