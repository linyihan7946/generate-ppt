export type SlideLayoutType = 'image_overlay' | 'image_only';
export type SlideImageSource = 'original' | 'ai_primary' | 'ai_fallback' | 'placeholder';
export type PlannerMode = 'strict' | 'creative';
export type DeckFormat = 'presenter' | 'detailed';
export type DeckAudience = 'general' | 'beginner' | 'executive' | 'student' | 'technical';
export type DeckFocus = 'overview' | 'timeline' | 'argument' | 'process' | 'comparison';
export type DeckStyle = 'professional' | 'minimal' | 'bold' | 'educational';
export type DeckLength = 'short' | 'default' | 'long';
export type SlideRole =
    | 'content'
    | 'agenda'
    | 'section_divider'
    | 'key_insight'
    | 'timeline'
    | 'comparison'
    | 'process'
    | 'data_highlight'
    | 'summary'
    | 'next_step';

export interface DeckBrief {
    deckGoal: string;
    audience: DeckAudience;
    focus: DeckFocus;
    style: DeckStyle;
    deckFormat: DeckFormat;
    desiredLength: DeckLength;
    chapterTitles: string[];
    coreTakeaways: string[];
}

export interface UnderstandingTopic {
    title: string;
    sourceRefs: number[];
    importance: number;
}

export interface UnderstandingResult {
    thesis: string;
    chapterTitles: string[];
    topics: UnderstandingTopic[];
    timelineSignals: string[];
    comparisonSignals: string[];
    processSignals: string[];
    keyNumbers: string[];
}

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
    slideRole?: SlideRole;
    keyMessage?: string;
    speakerNotes?: string[];
    sourceRefs?: number[];
}

export interface DocumentData {
    title: string;
    slides: SlideContent[];
    brief?: DeckBrief;
    understanding?: UnderstandingResult;
}

export interface PlannerOptions {
    mode?: PlannerMode;
    deckFormat?: DeckFormat;
    audience?: DeckAudience;
    focus?: DeckFocus;
    style?: DeckStyle;
    length?: DeckLength;
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
    slidesWithSummaryCount: number;
    slidesWithPromptCount: number;
    imageCoverage: number;
    summaryCoverage: number;
    promptCoverage: number;
    avgBulletsPerSlide: number;
    avgTextLengthPerSlide: number;
    avgBulletLength: number;
    levelJumpViolations: number;
    duplicateTitleCount: number;
    genericTitleCount: number;
    weakTransitionCount: number;
    actionCueSlideCount: number;
    redundantContentSlideCount: number;
    redundantContentItemCount: number;
    sparseContentSlideCount: number;
    severeSparseContentSlideCount: number;
    overlaySlideCount: number;
    imageOnlySlideCount: number;
    dominantLayoutRatio: number;
    overflowRiskSlideCount: number;
    promptAlignmentAvg: number;
    fallbackImageCount: number;
    renderedSlideCount: number;
    renderedSlideWithImageCount: number;
    renderedImageCoverage: number;
    renderedTextCoverage: number;
    renderedImageOnlySlideCount: number;
    renderedUniqueImageCount: number;
    renderedMetaArtifactSlideCount: number;
    renderedInstructionalTextSlideCount: number;
    renderedMixedLanguageSlideCount: number;
    visualFirstDeck: boolean;
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
        contentRichness: QualityDimensionScore;
        audienceFit: QualityDimensionScore;
        consistency: QualityDimensionScore;
    };
    metrics: QualityMetrics;
    keyFindings: string[];
    nextActions: string[];
}
