import { DocumentData, SlideContent, DeckBrief } from '../types';

/**
 * 将 slide 数据渲染为独立 HTML 页面（每页一个），用于后续截图
 * 设计分辨率: 1920×1080 (16:9)，渲染时使用 2x deviceScaleFactor 得到 3840×2160 高清图
 */
export class SlideRendererService {
    private readonly WIDTH = 1920;
    private readonly HEIGHT = 1080;

    /**
     * 为整个 deck 生成所有幻灯片的 HTML 字符串数组
     */
    renderAll(data: DocumentData): string[] {
        const slides = data.slides;
        const pages: string[] = [];

        // 封面
        pages.push(this.renderTitleSlide(data));

        // 内容页
        slides.forEach((slide, idx) => {
            const role = slide.slideRole || 'content';
            switch (role) {
                case 'agenda':
                    pages.push(this.renderAgendaSlide(slide, data.brief));
                    break;
                case 'comparison':
                    pages.push(this.renderComparisonSlide(slide));
                    break;
                case 'timeline':
                    pages.push(this.renderTimelineSlide(slide));
                    break;
                case 'summary':
                case 'next_step':
                    pages.push(this.renderSummarySlide(slide, role));
                    break;
                default:
                    pages.push(this.renderContentSlide(slide, idx, slides.length));
                    break;
            }
        });

        return pages;
    }

    private wrapPage(bodyContent: string, extraStyles: string = ''): string {
        return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;700;900&display=swap');
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    width: ${this.WIDTH}px;
    height: ${this.HEIGHT}px;
    overflow: hidden;
    font-family: 'Microsoft YaHei', 'PingFang SC', 'Noto Sans SC', sans-serif;
    -webkit-font-smoothing: antialiased;
  }
  .slide {
    width: ${this.WIDTH}px;
    height: ${this.HEIGHT}px;
    position: relative;
    overflow: hidden;
  }
  ${extraStyles}
</style>
</head>
<body>${bodyContent}</body>
</html>`;
    }

    private escapeHtml(text: string): string {
        return text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    // ======== Title Slide (封面) ========
    private renderTitleSlide(data: DocumentData): string {
        const brief = data.brief;
        const goal = brief?.deckGoal || '';
        const style = brief?.style || '';
        const audience = brief?.audience || '';

        const body = `
<div class="slide title-slide">
  <div class="title-bg">
    <div class="title-pattern"></div>
    <div class="title-gradient"></div>
  </div>
  <div class="title-content">
    <div class="title-badge">PRESENTATION</div>
    <h1 class="title-main">${this.escapeHtml(data.title)}</h1>
    ${goal ? `<p class="title-goal">${this.escapeHtml(goal)}</p>` : ''}
    <div class="title-meta">
      ${audience ? `<span class="meta-tag"><span class="meta-icon">👥</span> ${this.escapeHtml(String(audience))}</span>` : ''}
      ${style ? `<span class="meta-tag"><span class="meta-icon">🎨</span> ${this.escapeHtml(String(style))}</span>` : ''}
    </div>
  </div>
  <div class="title-decoration">
    <div class="deco-circle deco-c1"></div>
    <div class="deco-circle deco-c2"></div>
    <div class="deco-line deco-l1"></div>
  </div>
</div>`;

        const styles = `
.title-slide { background: #0a0f1e; color: #fff; }
.title-bg { position: absolute; inset: 0; }
.title-pattern {
  position: absolute; inset: 0;
  background-image:
    radial-gradient(circle at 20% 50%, rgba(14,165,233,0.15) 0%, transparent 50%),
    radial-gradient(circle at 80% 20%, rgba(139,92,246,0.12) 0%, transparent 40%);
}
.title-gradient {
  position: absolute; bottom: 0; left: 0; right: 0; height: 40%;
  background: linear-gradient(to top, rgba(10,15,30,0.9), transparent);
}
.title-content {
  position: relative; z-index: 2;
  padding: 120px 100px 0;
}
.title-badge {
  display: inline-block;
  padding: 6px 20px;
  font-size: 14px; font-weight: 700; letter-spacing: 3px;
  color: #0ea5e9;
  border: 2px solid rgba(14,165,233,0.4);
  border-radius: 30px;
  margin-bottom: 36px;
}
.title-main {
  font-size: 64px; font-weight: 900; line-height: 1.15;
  max-width: 1100px;
  background: linear-gradient(135deg, #fff 0%, #93c5fd 100%);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  margin-bottom: 28px;
}
.title-goal {
  font-size: 24px; color: rgba(255,255,255,0.7);
  max-width: 900px; line-height: 1.5;
  margin-bottom: 40px;
}
.title-meta { display: flex; gap: 24px; }
.meta-tag {
  display: flex; align-items: center; gap: 8px;
  padding: 8px 20px;
  background: rgba(255,255,255,0.06);
  border: 1px solid rgba(255,255,255,0.1);
  border-radius: 8px;
  font-size: 16px; color: rgba(255,255,255,0.6);
}
.meta-icon { font-size: 18px; }
.title-decoration { position: absolute; inset: 0; z-index: 1; pointer-events: none; }
.deco-circle {
  position: absolute; border-radius: 50%;
  border: 1px solid rgba(14,165,233,0.15);
}
.deco-c1 { width: 500px; height: 500px; right: -100px; top: -100px; }
.deco-c2 { width: 300px; height: 300px; right: 50px; top: 50px; border-color: rgba(139,92,246,0.12); }
.deco-line {
  position: absolute; width: 200px; height: 2px;
  background: linear-gradient(90deg, #0ea5e9, transparent);
  bottom: 100px; left: 100px;
}`;

        return this.wrapPage(body, styles);
    }

    // ======== Content Slide ========
    private renderContentSlide(slide: SlideContent, index: number, total: number): string {
        const keyMsg = slide.keyMessage || '';
        const bullets = slide.bullets || [];
        // 交替色调
        const accentHue = (index * 30 + 200) % 360;

        const bulletsHtml = bullets.map((b, i) => `
      <div class="bullet-item" style="animation-delay: ${i * 0.05}s">
        <div class="bullet-marker" style="background: hsl(${accentHue}, 70%, 55%);">${i + 1}</div>
        <div class="bullet-text">${this.escapeHtml(b)}</div>
      </div>`).join('');

        const hasImage = slide.images && slide.images.length > 0;
        const imageSection = hasImage
            ? `<div class="content-image">
                <img src="${slide.images[0]}" alt="" />
                <div class="image-overlay"></div>
               </div>`
            : `<div class="content-accent-block" style="background: linear-gradient(135deg, hsl(${accentHue},70%,95%), hsl(${accentHue},60%,88%));">
                <div class="accent-icon" style="color: hsl(${accentHue},70%,45%);">✦</div>
               </div>`;

        const body = `
<div class="slide content-slide">
  <div class="content-topbar">
    <div class="topbar-accent" style="background: hsl(${accentHue}, 70%, 55%);"></div>
    <span class="page-num">${String(index + 1).padStart(2, '0')} / ${String(total).padStart(2, '0')}</span>
  </div>
  <div class="content-layout ${hasImage ? 'has-image' : 'no-image'}">
    <div class="content-left">
      <h2 class="content-title">${this.escapeHtml(slide.title)}</h2>
      ${keyMsg ? `<p class="content-keymsg">${this.escapeHtml(keyMsg)}</p>` : ''}
      <div class="bullets-list">${bulletsHtml}</div>
    </div>
    <div class="content-right">
      ${imageSection}
    </div>
  </div>
</div>`;

        const styles = `
.content-slide { background: #fafbfe; color: #1a1e2e; }
.content-topbar {
  position: absolute; top: 0; left: 0; right: 0;
  height: 6px; background: #e8ecf2;
}
.topbar-accent { height: 100%; width: 30%; border-radius: 0 0 4px 0; }
.page-num {
  position: absolute; top: 20px; right: 60px;
  font-size: 14px; color: #94a3b8; font-weight: 500; letter-spacing: 1px;
}
.content-layout {
  display: flex; height: 100%; padding: 70px 60px 50px;
  gap: 50px;
}
.content-layout.has-image .content-left { flex: 1.1; }
.content-layout.has-image .content-right { flex: 0.9; display: flex; align-items: center; }
.content-layout.no-image .content-left { flex: 1.2; }
.content-layout.no-image .content-right { flex: 0.5; display: flex; align-items: center; justify-content: center; }
.content-title {
  font-size: 38px; font-weight: 800; color: #0f172a;
  margin-bottom: 16px; line-height: 1.25;
}
.content-keymsg {
  font-size: 18px; color: #64748b; margin-bottom: 30px;
  padding-left: 16px;
  border-left: 3px solid #cbd5e1;
  line-height: 1.5;
}
.bullets-list { display: flex; flex-direction: column; gap: 16px; }
.bullet-item {
  display: flex; align-items: flex-start; gap: 16px;
  padding: 14px 18px;
  background: #fff;
  border-radius: 12px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.04);
  border: 1px solid #f1f5f9;
}
.bullet-marker {
  flex-shrink: 0;
  width: 32px; height: 32px;
  display: flex; align-items: center; justify-content: center;
  border-radius: 8px;
  color: #fff; font-size: 14px; font-weight: 700;
}
.bullet-text { font-size: 18px; color: #334155; line-height: 1.55; padding-top: 3px; }
.content-image {
  width: 100%; height: 100%;
  border-radius: 16px; overflow: hidden;
  position: relative;
}
.content-image img {
  width: 100%; height: 100%; object-fit: cover;
}
.image-overlay {
  position: absolute; inset: 0;
  background: linear-gradient(to top, rgba(0,0,0,0.05), transparent);
}
.content-accent-block {
  width: 280px; height: 280px;
  border-radius: 24px;
  display: flex; align-items: center; justify-content: center;
}
.accent-icon { font-size: 80px; opacity: 0.4; }
`;

        return this.wrapPage(body, styles);
    }

    // ======== Agenda Slide ========
    private renderAgendaSlide(slide: SlideContent, brief?: DeckBrief): string {
        const items = (brief?.chapterTitles?.length ? brief.chapterTitles : slide.bullets).slice(0, 8);

        const itemsHtml = items.map((item, i) => `
      <div class="agenda-item">
        <div class="agenda-num">${String(i + 1).padStart(2, '0')}</div>
        <div class="agenda-label">${this.escapeHtml(item)}</div>
      </div>`).join('');

        const body = `
<div class="slide agenda-slide">
  <div class="agenda-bg"></div>
  <div class="agenda-content">
    <div class="agenda-header">
      <div class="agenda-badge">CONTENTS</div>
      <h2>${this.escapeHtml(slide.title)}</h2>
    </div>
    <div class="agenda-grid">${itemsHtml}</div>
  </div>
</div>`;

        const styles = `
.agenda-slide { background: #0f172a; color: #fff; }
.agenda-bg {
  position: absolute; inset: 0;
  background:
    radial-gradient(ellipse at 0% 100%, rgba(14,165,233,0.08) 0%, transparent 60%),
    radial-gradient(ellipse at 100% 0%, rgba(139,92,246,0.06) 0%, transparent 50%);
}
.agenda-content {
  position: relative; z-index: 2;
  padding: 80px 100px;
  height: 100%; display: flex; flex-direction: column;
}
.agenda-header { margin-bottom: 50px; }
.agenda-badge {
  display: inline-block;
  font-size: 13px; font-weight: 700; letter-spacing: 3px;
  color: #0ea5e9; margin-bottom: 16px;
}
.agenda-header h2 { font-size: 42px; font-weight: 800; }
.agenda-grid {
  display: grid; grid-template-columns: 1fr 1fr;
  gap: 20px; flex: 1;
}
.agenda-item {
  display: flex; align-items: center; gap: 20px;
  padding: 20px 28px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
  border-radius: 14px;
  transition: background 0.2s;
}
.agenda-num {
  font-size: 28px; font-weight: 900;
  color: #0ea5e9; letter-spacing: -1px;
  min-width: 50px;
}
.agenda-label { font-size: 20px; color: rgba(255,255,255,0.85); line-height: 1.4; }
`;

        return this.wrapPage(body, styles);
    }

    // ======== Comparison Slide ========
    private renderComparisonSlide(slide: SlideContent): string {
        const bullets = slide.bullets || [];
        const mid = Math.ceil(bullets.length / 2);
        const left = bullets.slice(0, mid);
        const right = bullets.slice(mid);

        const renderCol = (items: string[], color: string) => items.map((b, i) => `
      <div class="cmp-item">
        <div class="cmp-dot" style="background: ${color};"></div>
        <span>${this.escapeHtml(b)}</span>
      </div>`).join('');

        const body = `
<div class="slide comparison-slide">
  <div class="cmp-header">
    <h2>${this.escapeHtml(slide.title)}</h2>
    ${slide.keyMessage ? `<p class="cmp-keymsg">${this.escapeHtml(slide.keyMessage)}</p>` : ''}
  </div>
  <div class="cmp-columns">
    <div class="cmp-col cmp-col-a">
      <div class="cmp-col-title">A</div>
      ${renderCol(left, '#0ea5e9')}
    </div>
    <div class="cmp-divider"></div>
    <div class="cmp-col cmp-col-b">
      <div class="cmp-col-title">B</div>
      ${renderCol(right, '#8b5cf6')}
    </div>
  </div>
</div>`;

        const styles = `
.comparison-slide { background: #fafbfe; color: #1a1e2e; padding: 70px 80px; }
.cmp-header { margin-bottom: 40px; }
.cmp-header h2 { font-size: 38px; font-weight: 800; }
.cmp-keymsg { font-size: 18px; color: #64748b; margin-top: 10px; }
.cmp-columns { display: flex; gap: 0; flex: 1; }
.cmp-col { flex: 1; padding: 30px; }
.cmp-col-title {
  font-size: 20px; font-weight: 800; margin-bottom: 24px;
  padding: 6px 16px; border-radius: 8px; display: inline-block;
}
.cmp-col-a .cmp-col-title { background: rgba(14,165,233,0.1); color: #0ea5e9; }
.cmp-col-b .cmp-col-title { background: rgba(139,92,246,0.1); color: #8b5cf6; }
.cmp-divider { width: 2px; background: #e2e8f0; margin: 0 10px; border-radius: 1px; }
.cmp-item { display: flex; align-items: flex-start; gap: 14px; margin-bottom: 18px; font-size: 18px; line-height: 1.55; }
.cmp-dot { width: 10px; height: 10px; border-radius: 50%; margin-top: 8px; flex-shrink: 0; }
`;

        return this.wrapPage(body, styles);
    }

    // ======== Timeline Slide ========
    private renderTimelineSlide(slide: SlideContent): string {
        const bullets = slide.bullets || [];

        const itemsHtml = bullets.map((b, i) => `
      <div class="tl-item">
        <div class="tl-dot"></div>
        <div class="tl-connector"></div>
        <div class="tl-card">
          <div class="tl-num">STEP ${i + 1}</div>
          <div class="tl-text">${this.escapeHtml(b)}</div>
        </div>
      </div>`).join('');

        const body = `
<div class="slide timeline-slide">
  <div class="tl-header">
    <h2>${this.escapeHtml(slide.title)}</h2>
  </div>
  <div class="tl-track">${itemsHtml}</div>
</div>`;

        const styles = `
.timeline-slide { background: #0f172a; color: #fff; padding: 70px 80px; }
.tl-header h2 { font-size: 38px; font-weight: 800; margin-bottom: 50px; }
.tl-track { display: flex; gap: 24px; align-items: flex-start; overflow: hidden; }
.tl-item { flex: 1; display: flex; flex-direction: column; align-items: center; position: relative; }
.tl-dot {
  width: 16px; height: 16px; border-radius: 50%;
  background: #0ea5e9; box-shadow: 0 0 20px rgba(14,165,233,0.4);
  z-index: 2;
}
.tl-connector {
  width: 100%; height: 2px; background: rgba(14,165,233,0.3);
  position: absolute; top: 7px; left: 50%;
}
.tl-item:last-child .tl-connector { display: none; }
.tl-card {
  margin-top: 20px;
  padding: 20px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
  border-radius: 12px;
  text-align: center;
  width: 100%;
}
.tl-num { font-size: 12px; font-weight: 700; color: #0ea5e9; letter-spacing: 2px; margin-bottom: 10px; }
.tl-text { font-size: 16px; color: rgba(255,255,255,0.8); line-height: 1.5; }
`;

        return this.wrapPage(body, styles);
    }

    // ======== Summary / Next Step Slide ========
    private renderSummarySlide(slide: SlideContent, role: string): string {
        const bullets = slide.bullets || [];
        const isSummary = role === 'summary';

        const itemsHtml = bullets.map((b, i) => `
      <div class="sum-item">
        <div class="sum-check">${isSummary ? '✓' : '→'}</div>
        <div class="sum-text">${this.escapeHtml(b)}</div>
      </div>`).join('');

        const body = `
<div class="slide summary-slide">
  <div class="sum-bg"></div>
  <div class="sum-content">
    <div class="sum-badge">${isSummary ? 'SUMMARY' : 'NEXT STEPS'}</div>
    <h2>${this.escapeHtml(slide.title)}</h2>
    ${slide.keyMessage ? `<p class="sum-keymsg">${this.escapeHtml(slide.keyMessage)}</p>` : ''}
    <div class="sum-list">${itemsHtml}</div>
  </div>
</div>`;

        const styles = `
.summary-slide { background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #fff; }
.sum-bg {
  position: absolute; inset: 0;
  background: radial-gradient(ellipse at 70% 30%, rgba(14,165,233,0.1) 0%, transparent 60%);
}
.sum-content { position: relative; z-index: 2; padding: 80px 100px; }
.sum-badge {
  font-size: 13px; font-weight: 700; letter-spacing: 3px;
  color: ${isSummary ? '#10b981' : '#f59e0b'}; margin-bottom: 20px;
}
.sum-content h2 { font-size: 42px; font-weight: 800; margin-bottom: 16px; }
.sum-keymsg { font-size: 20px; color: rgba(255,255,255,0.6); margin-bottom: 40px; }
.sum-list { display: flex; flex-direction: column; gap: 16px; }
.sum-item {
  display: flex; align-items: flex-start; gap: 18px;
  padding: 18px 24px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.08);
  border-radius: 12px;
}
.sum-check {
  font-size: 20px; font-weight: 700;
  color: ${isSummary ? '#10b981' : '#f59e0b'};
  min-width: 28px; text-align: center;
}
.sum-text { font-size: 18px; color: rgba(255,255,255,0.85); line-height: 1.55; }
`;

        return this.wrapPage(body, styles);
    }
}
