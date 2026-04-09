# NotebookLM-Style PPT Redesign Plan

## Goal

Upgrade the current pipeline from:

`Parse -> Plan slides -> Generate images -> Render PPT -> Evaluate`

to:

`Parse -> Understand -> Brief -> Plan -> Render -> Critique -> Revise`

The target is not just "generate a valid PPT", but to generate a deck that:

- reads like a presentation instead of a source document dump
- adapts to audience, length, and presentation mode
- uses different slide archetypes based on content shape
- keeps claims grounded in source material
- can self-criticize and repair weak slides

## Current Gaps

### 1. Parsing is too slide-oriented too early

Current parsing already outputs `DocumentData.slides`, which makes source structure become PPT structure by default.

Impact:

- the system preserves document hierarchy, but does not reframe content into a presentation narrative
- weak source sections become weak slides
- long sections become split pages instead of better page types

Relevant code:

- `src/services/parser.service.ts`

### 2. Planning is local, not deck-level

The planner currently asks for per-slide JSON and then does sparse-page expansion.

Impact:

- good at bullet cleanup
- weak at global storyline, audience adaptation, and page rhythm
- no clear deck thesis, chapter flow, or intentional climax

Relevant code:

- `src/services/planner.service.ts`

### 3. Rendering supports too few visual archetypes

The renderer is mainly:

- title slide
- image overlay slide
- image-only slide

Impact:

- many pages look visually similar
- timelines, comparisons, key-number slides, and section breaks are not first-class
- output feels "AI wallpaper + bullets" rather than presentation-native

Relevant code:

- `src/services/ppt.service.ts`

### 4. Evaluation is post-hoc, not revision-driving

The evaluator already gives meaningful scores, but generation does not yet use those scores to repair weak slides automatically.

Relevant code:

- `src/services/evaluator.service.ts`

## Target Product Behavior

The redesigned system should behave closer to NotebookLM's strengths:

- source-grounded synthesis instead of direct source projection
- explicit control over audience, focus, style, and deck format
- support for "presenter slides" vs "detailed deck"
- multiple page archetypes instead of one default template
- iterative revision of weak pages

## Proposed Architecture

### Stage 1. Parse

Responsibility:

- read markdown / docx / pdf
- preserve hierarchy
- retain source chunks, not just future slides

New output should be source-centric, for example:

```ts
interface SourceChunk {
  id: string;
  text: string;
  level?: number;
  heading?: string;
  breadcrumb?: string;
  sourceType: 'paragraph' | 'list_item' | 'heading' | 'table' | 'image_caption';
  order: number;
}
```

Instead of directly mapping everything into `SlideContent`, parsing should produce a richer intermediate document.

Recommendation:

- keep `DocumentData` for compatibility
- add a richer internal model such as `ParsedDocument`

### Stage 2. Understand

Responsibility:

- turn raw chunks into structured understanding
- extract deck-worthy concepts

New service:

- `src/services/understanding.service.ts`

Core outputs:

- key topics
- central thesis
- entities
- dates
- numbers
- comparisons
- process steps
- timeline events
- quotes / claims
- evidence links back to source chunks

Suggested model:

```ts
interface UnderstandingResult {
  thesis: string;
  topics: TopicNode[];
  insights: InsightItem[];
  timeline: TimelineEvent[];
  comparisons: ComparisonItem[];
  keyNumbers: KeyNumberItem[];
  sourceMap: Record<string, string[]>;
}
```

This is the biggest capability gap today.

### Stage 3. Brief

Responsibility:

- define what kind of deck to make before planning slides

New service:

- `src/services/brief.service.ts`

Inputs:

- `UnderstandingResult`
- user preferences

Suggested inputs:

- `audience`
- `focus`
- `deckFormat`
- `length`
- `language`
- `style`

Suggested output:

```ts
interface DeckBrief {
  deckGoal: string;
  audience: string;
  style: string;
  deckFormat: 'presenter' | 'detailed';
  desiredLength: 'short' | 'default' | 'long';
  coreTakeaways: string[];
  chapterPlan: ChapterPlan[];
}
```

This is where we close the gap with NotebookLM's "choose format / audience / focus / length" behavior.

### Stage 4. Plan

Responsibility:

- convert `DeckBrief + UnderstandingResult` into slides
- assign slide roles and content intent

Refactor existing planner into:

- `NarrativePlanner`
- `SparseExpansionPlanner`
- `RevisionPlanner`

Each slide should gain a `slideRole`:

- `cover`
- `agenda`
- `section-divider`
- `key-insight`
- `timeline`
- `comparison`
- `process`
- `quote`
- `case-study`
- `data-highlight`
- `summary`
- `next-step`

Suggested slide shape:

```ts
interface PlannedSlideV2 {
  id: string;
  slideRole: SlideRole;
  title: string;
  keyMessage: string;
  bullets: string[];
  speakerNotes?: string[];
  sourceRefs: string[];
  visualSpec: VisualSpec;
  layoutSpec: LayoutSpec;
}
```

Key rule:

- the planner should be allowed to merge, split, or reorder source sections when deck quality improves
- only source grounding must remain strict

That is a deliberate change from the current "must preserve source slide order" behavior.

### Stage 5. Render

Responsibility:

- map slide roles into visual archetypes

Refactor `ppt.service.ts` from one-template rendering into archetype-based rendering.

Minimum archetypes:

- cover
- agenda
- section divider
- hero insight
- bullets with visual
- comparison two-column
- timeline
- process flow
- quote / evidence
- summary / next step

Suggested file split:

- `src/services/ppt/archetypes/cover.ts`
- `src/services/ppt/archetypes/timeline.ts`
- `src/services/ppt/archetypes/comparison.ts`
- `src/services/ppt/archetypes/summary.ts`
- `src/services/ppt/ppt-renderer.service.ts`

Visual rule:

- not every page should use a generated background image
- diagrams, numbers, comparisons, and section pages often work better without AI wallpaper

### Stage 6. Critique

Responsibility:

- review generated deck as if it were a human design / content reviewer

Critique dimensions:

- story clarity
- deck pacing
- slide role appropriateness
- redundancy
- source grounding
- readability
- visual rhythm

New output:

```ts
interface DeckCritique {
  overallVerdict: string;
  slideIssues: SlideIssue[];
  revisionTasks: RevisionTask[];
}
```

### Stage 7. Revise

Responsibility:

- revise only weak slides
- do not regenerate the whole deck when only 2-3 slides are poor

Revision examples:

- convert bullet page to timeline
- shorten overloaded slide
- expand sparse slide using neighbor context
- add missing conclusion
- improve section transition

## Data Model Changes

### Keep for compatibility

- `SlideContent`
- `DocumentData`

### Add new types

- `ParsedDocument`
- `SourceChunk`
- `UnderstandingResult`
- `DeckBrief`
- `PlannedSlideV2`
- `SlideRole`
- `VisualSpec`
- `LayoutSpec`
- `DeckCritique`
- `RevisionTask`

Recommendation:

- add these in `src/types.ts`
- keep old interfaces until migration is complete

## Concrete Refactor Plan

### Phase 1. Introduce deck brief and slide roles

Target:

- biggest quality gain with moderate change cost

Tasks:

- add `DeckFormat`, `SlideRole`, `DeckBrief`, `PlannedSlideV2` types
- add planner inputs for `audience`, `focus`, `style`, `length`, `deckFormat`
- make planner produce `slideRole`
- stop treating every page as generic bullet page

Files:

- `src/types.ts`
- `src/services/planner.service.ts`
- `src/index.ts`
- `src/cli.ts`

Expected gain:

- better page intent
- better page rhythm
- closer to NotebookLM "Presenter vs Detailed" behavior

### Phase 2. Expand renderer archetypes

Tasks:

- replace binary layout logic with slide-role-based rendering
- add at least 4 new archetypes: `timeline`, `comparison`, `section-divider`, `summary`
- reduce image dependency on slides that should be diagrammatic

Files:

- `src/services/ppt.service.ts`

Expected gain:

- the deck starts feeling designed, not just generated

### Phase 3. Add understanding layer

Tasks:

- create `understanding.service.ts`
- extract thesis, key themes, dates, comparisons, and important facts
- bind bullets back to source chunk IDs

Files:

- `src/services/parser.service.ts`
- `src/services/understanding.service.ts`
- `src/types.ts`

Expected gain:

- much stronger content restructuring
- less dependence on raw heading hierarchy

### Phase 4. Add critique + revise loop

Tasks:

- evaluate at slide level, not only deck level
- generate revision instructions for weak slides
- rerun planner only for weak slides

Files:

- `src/services/evaluator.service.ts`
- `src/services/revision.service.ts`
- `src/index.ts`
- `src/cli.ts`

Expected gain:

- stable deck quality
- fewer obviously weak pages

## Product-Level Controls To Add

Recommended new API / CLI options:

- `deckFormat=presenter|detailed`
- `audience=beginner|executive|student|technical`
- `focus=overview|timeline|argument|process|comparison`
- `length=short|default|long`
- `style=professional|minimal|bold|educational`
- `language=zh|en`

API example:

```http
POST /generate-ppt
form-data:
  file=@input.docx
  plannerMode=creative
  deckFormat=presenter
  audience=beginner
  focus=overview
  length=default
  style=professional
```

## Prompt Strategy Changes

Current planner prompt is too slide-local.

Change it to a chain:

1. `understand sources`
2. `build deck brief`
3. `build chapter plan`
4. `plan slides with slide roles`
5. `revise weak slides`

Important prompt rules:

- allow reorganization for presentation quality
- preserve source grounding
- prefer fewer stronger slides over one-source-node-one-slide
- create explicit transitions between sections
- write page text differently for `presenter` and `detailed`

## Evaluation Upgrade

Your evaluation system is already improved, but to support the new architecture it should add:

- `slideRoleCoverage`
- `roleDiversity`
- `sectionTransitionQuality`
- `sourceGroundingCoverage`
- `citationCoverage`
- `notesCoverage`
- `repairSuccessRate`

And most importantly:

- evaluation should output revision tasks, not just scores

## Acceptance Criteria

The redesign can be considered successful when:

- the same source document produces clearly different decks for `presenter` vs `detailed`
- at least 5 slide archetypes are used in medium-length decks
- sparse pages are expanded using section context rather than generic filler
- most bullets can be traced to source chunks
- generated decks no longer look like one-template image overlays
- evaluation can trigger targeted slide repair automatically

## Suggested Build Order

Recommended implementation order:

1. Phase 1: brief + slide roles
2. Phase 2: renderer archetypes
3. Phase 3: understanding layer
4. Phase 4: critique + revise loop

This order gives the best quality gain per engineering effort.

## Recommendation

If only one iteration is possible in the near term, do this:

- add `deckFormat`
- add `audience / focus / length / style`
- add `slideRole`
- implement 4 new renderer archetypes

That alone will make the output noticeably closer to NotebookLM than only tuning prompts or switching models.
