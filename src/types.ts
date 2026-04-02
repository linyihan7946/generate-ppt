export type SlideLayoutType = 'image_overlay' | 'image_only';
export type SlideImageSource = 'original' | 'ai_primary' | 'ai_fallback' | 'placeholder';
export type PlannerMode = 'strict' | 'creative';

export interface SlideContent {
    title: string;
    bullets: string[];
    images: string[]; // Base64 or URLs of extracted/generated images
    level?: number;
    breadcrumb?: string;
    summary?: string;
    layout?: SlideLayoutType;
    imageIntent?: string;
    imagePrompt?: string;
    sourceIndex?: number;
    imageSource?: SlideImageSource;
}

export interface DocumentData {
    title: string;
    slides: SlideContent[];
}

export interface PlannerOptions {
    mode?: PlannerMode;
}

export interface QualityDimensionScore {
    name: string;
    score: number;
    weight: number;
    weightedScore: number;
    evidence: string[];
    issues: string[];
    recommendations: string[];
}

export interface QualityMetrics {
    slideCount: number;
    slideWithImageCount: number;
    imageCoverage: number;
    avgBulletsPerSlide: number;
    avgBulletLength: number;
    levelJumpViolations: number;
    duplicateTitleCount: number;
    redundantContentSlideCount: number;
    redundantContentItemCount: number;
    overlaySlideCount: number;
    imageOnlySlideCount: number;
    overflowRiskSlideCount: number;
    promptAlignmentAvg: number;
    fallbackImageCount: number;
}

export interface QualityReport {
    version: string;
    generatedAt: string;
    title: string;
    outputPath?: string;
    overallScore: number;
    grade: string;
    dimensions: {
        logic: QualityDimensionScore;
        layout: QualityDimensionScore;
        imageSemantics: QualityDimensionScore;
    };
    metrics: QualityMetrics;
    keyFindings: string[];
    nextActions: string[];
}
