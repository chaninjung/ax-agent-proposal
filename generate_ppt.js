const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icon imports
const { FaUsers, FaExclamationTriangle, FaLightbulb, FaCogs, FaFileAlt, FaGraduationCap, FaChartLine, FaRocket, FaQuestionCircle, FaDatabase, FaRobot, FaSlack, FaArrowRight, FaCheck, FaTimes, FaUserTie, FaCode, FaBullhorn, FaEllipsisH, FaExpandArrowsAlt, FaClock, FaBuilding } = require("react-icons/fa");
const { SiConfluence } = require("react-icons/si");

// Colors - White background + Carrot Global CI (orange/carrot)
const CARROT_ORANGE = "FF6B2C";       // Primary CI color
const CARROT_DARK = "E55A1B";         // Darker accent
const DARK_TEXT = "2D2D2D";           // Main text
const MEDIUM_TEXT = "555555";         // Secondary text
const LIGHT_GRAY = "999999";          // Muted text
const WHITE = "FFFFFF";
const LIGHT_BG = "FFFFFF";            // White background for all content slides
const CARD_BG = "F7F7F7";            // Subtle gray cards
const SOFT_ORANGE_BG = "FFF5EE";     // Soft orange tint
const ACCENT_BLUE = "2B6CB0";        // Secondary accent for contrast
const ACCENT_GREEN = "38A169";
const ACCENT_RED = "E53E3E";
const HEADER_BG = "2D2D2D";          // Dark header bar

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

const makeShadow = () => ({ type: "outer", blur: 4, offset: 1, angle: 135, color: "000000", opacity: 0.08 });

async function main() {
  let pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "정찬인";
  pres.title = "AX 팀 업무 효율화를 위한 AI 에이전트 기반 소스자료 공유 시스템";

  // Pre-render icons
  const iconUsersWhite = await iconToBase64Png(FaUsers, "#FFFFFF");
  const iconWarningWhite = await iconToBase64Png(FaExclamationTriangle, "#FFFFFF");
  const iconLightbulbWhite = await iconToBase64Png(FaLightbulb, "#FFFFFF");
  const iconCogsWhite = await iconToBase64Png(FaCogs, "#FFFFFF");
  const iconFileAltWhite = await iconToBase64Png(FaFileAlt, "#FFFFFF");
  const iconGradWhite = await iconToBase64Png(FaGraduationCap, "#FFFFFF");
  const iconChartWhite = await iconToBase64Png(FaChartLine, "#FFFFFF");
  const iconRocketWhite = await iconToBase64Png(FaRocket, "#FFFFFF");

  const iconArrow = await iconToBase64Png(FaArrowRight, "#" + CARROT_ORANGE);
  const iconUserTie = await iconToBase64Png(FaUserTie, "#" + CARROT_ORANGE);
  const iconCode = await iconToBase64Png(FaCode, "#" + CARROT_ORANGE);
  const iconBullhorn = await iconToBase64Png(FaBullhorn, "#" + CARROT_ORANGE);
  const iconEllipsis = await iconToBase64Png(FaEllipsisH, "#" + LIGHT_GRAY);
  const iconBuilding = await iconToBase64Png(FaBuilding, "#" + CARROT_ORANGE);
  const iconDatabase = await iconToBase64Png(FaDatabase, "#FFFFFF");
  const iconRobot = await iconToBase64Png(FaRobot, "#FFFFFF");
  const iconFileAltOrange = await iconToBase64Png(FaFileAlt, "#FFFFFF");

  const iconConfluence = await iconToBase64Png(SiConfluence, "#" + ACCENT_BLUE);
  const iconRobotOrange = await iconToBase64Png(FaRobot, "#" + CARROT_ORANGE);
  const iconCogsOrange = await iconToBase64Png(FaCogs, "#" + CARROT_ORANGE);
  const iconSlackOrange = await iconToBase64Png(FaSlack, "#" + CARROT_ORANGE);

  const iconChartGreen = await iconToBase64Png(FaChartLine, "#" + ACCENT_GREEN);
  const iconExpandBlue = await iconToBase64Png(FaExpandArrowsAlt, "#" + ACCENT_BLUE);
  const iconUsersOrange = await iconToBase64Png(FaUsers, "#" + CARROT_ORANGE);
  const iconRocketOrange = await iconToBase64Png(FaRocket, "#" + CARROT_ORANGE);

  // ===== SLIDE 1: TITLE =====
  let s1 = pres.addSlide();
  s1.background = { color: WHITE };
  // Top accent bar
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: CARROT_ORANGE } });
  // Left accent block
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: CARROT_ORANGE } });

  s1.addText("AX 팀 업무 효율화를 위한", {
    x: 0.8, y: 1.6, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: "Malgun Gothic", color: CARROT_ORANGE, bold: true,
    align: "left", margin: 0
  });
  s1.addText("AI 에이전트 기반\n소스자료 공유 시스템", {
    x: 0.8, y: 2.1, w: 8.4, h: 1.6,
    fontSize: 38, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true,
    align: "left", margin: 0
  });
  s1.addText("캐럿글로벌 AX사업부 실무 적용 과제", {
    x: 0.8, y: 3.8, w: 8.4, h: 0.5,
    fontSize: 14, fontFace: "Malgun Gothic", color: LIGHT_GRAY,
    align: "left", margin: 0
  });

  // Bottom section
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.725, fill: { color: CARROT_ORANGE } });
  s1.addText("발표자: 정찬인", {
    x: 0.8, y: 4.9, w: 4, h: 0.725,
    fontSize: 14, fontFace: "Malgun Gothic", color: WHITE, bold: false,
    align: "left", valign: "middle", margin: 0
  });

  // ===== SLIDE 2: AX TEAM STATUS =====
  let s2 = pres.addSlide();
  s2.background = { color: WHITE };
  // Header
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: CARROT_ORANGE } });
  s2.addImage({ data: iconUsersWhite, x: 0.5, y: 0.17, w: 0.48, h: 0.48 });
  s2.addText("캐럿글로벌 AX 팀 현황", {
    x: 1.1, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true,
    align: "left", valign: "middle", margin: 0
  });

  // Team composition card
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 4.2, h: 3.9, fill: { color: CARD_BG }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 0.08, h: 3.9, fill: { color: CARROT_ORANGE } });
  s2.addText("현재 팀 구성 (6명)", {
    x: 0.85, y: 1.3, w: 3.6, h: 0.4,
    fontSize: 16, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0
  });

  // Team member icons
  const teamMembers = [
    { icon: iconUserTie, label: "팀리더 1명", x: 0.85 },
    { icon: iconBullhorn, label: "마케터 1명", x: 2.05 },
    { icon: iconCode, label: "AI개발 1명", x: 3.25 },
  ];
  for (const m of teamMembers) {
    s2.addImage({ data: m.icon, x: m.x + 0.15, y: 1.9, w: 0.4, h: 0.4 });
    s2.addText(m.label, { x: m.x - 0.1, y: 2.35, w: 1.0, h: 0.35, fontSize: 10, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, align: "center", margin: 0 });
  }
  s2.addImage({ data: iconEllipsis, x: 4.1, y: 2.0, w: 0.3, h: 0.3 });
  s2.addText("+ 기타 3명", { x: 3.75, y: 2.35, w: 1.0, h: 0.35, fontSize: 10, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, align: "center", margin: 0 });

  s2.addText([
    { text: "AI개발 1명 + 트레이니/계약직 협업 구조", options: { bullet: true, breakLine: true } },
    { text: "매일 아침 10분 스탠드업 미팅", options: { bullet: true } },
  ], { x: 0.85, y: 3.0, w: 3.6, h: 1.5, fontSize: 12, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0, paraSpaceAfter: 6 });

  // Right cards
  s2.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.15, w: 4.2, h: 1.7, fill: { color: CARD_BG }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.15, w: 0.08, h: 1.7, fill: { color: CARROT_ORANGE } });
  s2.addText("주요 고객사", {
    x: 5.65, y: 1.3, w: 3.6, h: 0.35,
    fontSize: 16, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0
  });
  s2.addImage({ data: iconBuilding, x: 5.65, y: 1.85, w: 0.3, h: 0.3 });
  s2.addText("대상  |  현대엔지니어링  |  GS리테일 등", {
    x: 6.05, y: 1.85, w: 3.2, h: 0.35,
    fontSize: 12, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0, valign: "middle"
  });
  s2.addText("다수 고객사 동시 운영 중", {
    x: 6.05, y: 2.2, w: 3.2, h: 0.3,
    fontSize: 11, fontFace: "Malgun Gothic", color: LIGHT_GRAY, margin: 0
  });

  // Expansion card
  s2.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.15, w: 4.2, h: 1.9, fill: { color: CARD_BG }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 3.15, w: 0.08, h: 1.9, fill: { color: CARROT_DARK } });
  s2.addText("확장 계획", {
    x: 5.65, y: 3.3, w: 3.6, h: 0.35,
    fontSize: 16, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0
  });
  s2.addText("6명", { x: 5.65, y: 3.8, w: 1.2, h: 0.6, fontSize: 30, fontFace: "Malgun Gothic", color: CARROT_ORANGE, bold: true, margin: 0, align: "center" });
  s2.addImage({ data: iconArrow, x: 6.95, y: 3.9, w: 0.35, h: 0.35 });
  s2.addText("12명", { x: 7.4, y: 3.8, w: 1.2, h: 0.6, fontSize: 30, fontFace: "Malgun Gothic", color: CARROT_DARK, bold: true, margin: 0, align: "center" });
  s2.addText("AI 인력 2~3명 + B2B/B2G 영업 채용 예정", {
    x: 5.65, y: 4.5, w: 3.6, h: 0.35,
    fontSize: 11, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0
  });

  // ===== SLIDE 3: PROBLEM DEFINITION =====
  let s3 = pres.addSlide();
  s3.background = { color: WHITE };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: CARROT_ORANGE } });
  s3.addImage({ data: iconWarningWhite, x: 0.5, y: 0.17, w: 0.48, h: 0.48 });
  s3.addText("문제 정의: 3가지 병목", {
    x: 1.1, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true,
    align: "left", valign: "middle", margin: 0
  });

  const problems = [
    { num: "01", title: "AI 전문 지식 집중", desc: "AI 관련 질문과 확인이\n1명에게 몰림\n\n팀 확장 시 병목 심화", x: 0.5 },
    { num: "02", title: "소스 자료 분산", desc: "과거 프로젝트 자료,\n산업별 레퍼런스가\n개인별로 흩어져 있음\n\n재활용 어려움", x: 3.5 },
    { num: "03", title: "산출물 반복 제작", desc: "고객사별 제안서,\n교육자료를 매번\n처음부터 작성\n\n시간과 품질 편차 발생", x: 6.5 },
  ];
  for (const p of problems) {
    s3.addShape(pres.shapes.RECTANGLE, { x: p.x, y: 1.15, w: 2.7, h: 3.9, fill: { color: CARD_BG }, shadow: makeShadow() });
    s3.addShape(pres.shapes.RECTANGLE, { x: p.x, y: 1.15, w: 0.08, h: 3.9, fill: { color: CARROT_ORANGE } });
    s3.addText(p.num, { x: p.x + 0.25, y: 1.35, w: 0.8, h: 0.5, fontSize: 30, fontFace: "Malgun Gothic", color: CARROT_ORANGE, bold: true, margin: 0 });
    s3.addText(p.title, { x: p.x + 0.25, y: 1.95, w: 2.2, h: 0.4, fontSize: 15, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0 });
    s3.addShape(pres.shapes.RECTANGLE, { x: p.x + 0.25, y: 2.4, w: 1.0, h: 0.03, fill: { color: CARROT_ORANGE } });
    s3.addText(p.desc, { x: p.x + 0.25, y: 2.6, w: 2.2, h: 2.2, fontSize: 12, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0 });
  }

  // ===== SLIDE 4: SOLUTION DIRECTION =====
  let s4 = pres.addSlide();
  s4.background = { color: WHITE };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: CARROT_ORANGE } });
  s4.addImage({ data: iconLightbulbWhite, x: 0.5, y: 0.17, w: 0.48, h: 0.48 });
  s4.addText("해결 방향", {
    x: 1.1, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true,
    align: "left", valign: "middle", margin: 0
  });

  // Core principle
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 9, h: 1.1, fill: { color: SOFT_ORANGE_BG }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 0.08, h: 1.1, fill: { color: CARROT_ORANGE } });
  s4.addText("핵심 원칙", {
    x: 0.85, y: 1.2, w: 2, h: 0.3,
    fontSize: 11, fontFace: "Malgun Gothic", color: CARROT_ORANGE, bold: true, margin: 0
  });
  s4.addText('"소스 자료는 공유, 산출물은 각자"', {
    x: 0.85, y: 1.55, w: 8.4, h: 0.55,
    fontSize: 24, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0,
    align: "center"
  });

  // Three columns
  const directions = [
    { title: "소스 자료 축적", desc: "타 회사 자료, 산업별\n레퍼런스, 과거 프로젝트\n데이터를 공유 저장소에\n체계적으로 축적", icon: iconDatabase },
    { title: "AI 자동화", desc: "AI 도구가 소스 자료를\n기반으로 산출물 초안을\n자동 생성\n— 반복 작업 제거", icon: iconRobot },
    { title: "개인 산출물", desc: "제안서, 교육 PPT,\n커리큘럼은 각자의 것\n— 자율성과 품질\n모두 확보", icon: iconFileAltOrange },
  ];
  for (let i = 0; i < directions.length; i++) {
    const d = directions[i];
    const x = 0.5 + i * 3.15;
    s4.addShape(pres.shapes.RECTANGLE, { x, y: 2.6, w: 2.85, h: 2.7, fill: { color: CARD_BG }, shadow: makeShadow() });
    // Icon circle
    s4.addShape(pres.shapes.OVAL, { x: x + 1.05, y: 2.8, w: 0.7, h: 0.7, fill: { color: CARROT_ORANGE } });
    s4.addImage({ data: d.icon, x: x + 1.15, y: 2.9, w: 0.5, h: 0.5 });
    s4.addText(d.title, { x: x + 0.2, y: 3.6, w: 2.45, h: 0.35, fontSize: 15, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, align: "center", margin: 0 });
    s4.addShape(pres.shapes.RECTANGLE, { x: x + 0.9, y: 3.98, w: 1.0, h: 0.03, fill: { color: CARROT_ORANGE } });
    s4.addText(d.desc, { x: x + 0.2, y: 4.1, w: 2.45, h: 1.1, fontSize: 11, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, align: "center", margin: 0 });
  }
  s4.addImage({ data: iconArrow, x: 3.2, y: 3.5, w: 0.3, h: 0.3 });
  s4.addImage({ data: iconArrow, x: 6.35, y: 3.5, w: 0.3, h: 0.3 });

  // ===== SLIDE 5: SYSTEM ARCHITECTURE =====
  let s5 = pres.addSlide();
  s5.background = { color: WHITE };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: CARROT_ORANGE } });
  s5.addImage({ data: iconCogsWhite, x: 0.5, y: 0.17, w: 0.48, h: 0.48 });
  s5.addText("시스템 구조", {
    x: 1.1, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true,
    align: "left", valign: "middle", margin: 0
  });

  const tools = [
    { name: "Confluence", role: "소스 자료 저장 · 관리", desc: "산업별 레퍼런스, 과거\n프로젝트 데이터 축적", icon: iconConfluence, x: 0.5, y: 1.15 },
    { name: "Claude", role: "AI 엔진", desc: "소스 자료 기반\n산출물 초안 생성", icon: iconRobotOrange, x: 5.3, y: 1.15 },
    { name: "n8n / Dify", role: "워크플로우 자동화", desc: "자료 수집 → AI 처리\n→ 산출물 출력", icon: iconCogsOrange, x: 0.5, y: 3.15 },
    { name: "Slack", role: "팀 소통 + 알림", desc: "프로젝트 채널,\n자동 알림, 공유", icon: iconSlackOrange, x: 5.3, y: 3.15 },
  ];
  for (const t of tools) {
    s5.addShape(pres.shapes.RECTANGLE, { x: t.x, y: t.y, w: 4.2, h: 1.7, fill: { color: CARD_BG }, shadow: makeShadow() });
    s5.addShape(pres.shapes.RECTANGLE, { x: t.x, y: t.y, w: 0.08, h: 1.7, fill: { color: CARROT_ORANGE } });
    s5.addImage({ data: t.icon, x: t.x + 0.3, y: t.y + 0.35, w: 0.5, h: 0.5 });
    s5.addText(t.name, { x: t.x + 1.0, y: t.y + 0.2, w: 2.9, h: 0.35, fontSize: 16, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0 });
    s5.addText(t.role, { x: t.x + 1.0, y: t.y + 0.5, w: 2.9, h: 0.3, fontSize: 11, fontFace: "Malgun Gothic", color: CARROT_ORANGE, bold: true, margin: 0 });
    s5.addText(t.desc, { x: t.x + 1.0, y: t.y + 0.85, w: 2.9, h: 0.65, fontSize: 11, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0 });
  }
  // Flow arrows
  s5.addImage({ data: iconArrow, x: 4.85, y: 1.8, w: 0.3, h: 0.3 });
  s5.addImage({ data: iconArrow, x: 4.85, y: 3.8, w: 0.3, h: 0.3 });
  s5.addShape(pres.shapes.RECTANGLE, { x: 2.55, y: 2.9, w: 0.04, h: 0.25, fill: { color: CARROT_ORANGE } });
  s5.addShape(pres.shapes.RECTANGLE, { x: 7.35, y: 2.9, w: 0.04, h: 0.25, fill: { color: CARROT_ORANGE } });

  // ===== SLIDE 6: SCENARIO 1 - PROPOSALS =====
  let s6 = pres.addSlide();
  s6.background = { color: WHITE };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: CARROT_ORANGE } });
  s6.addImage({ data: iconFileAltWhite, x: 0.5, y: 0.17, w: 0.48, h: 0.48 });
  s6.addText("적용 시나리오 ① 제안서 자동화", {
    x: 1.1, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true,
    align: "left", valign: "middle", margin: 0
  });

  // Before
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 4.2, h: 3.2, fill: { color: CARD_BG }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 0.08, h: 3.2, fill: { color: ACCENT_RED } });
  s6.addText("Before", { x: 0.85, y: 1.3, w: 1.5, h: 0.4, fontSize: 18, fontFace: "Malgun Gothic", color: ACCENT_RED, bold: true, margin: 0 });
  s6.addText([
    { text: "매번 맨땅에서 제안서 시작", options: { bullet: true, breakLine: true } },
    { text: "AI담당자에게 매번 확인 필요", options: { bullet: true, breakLine: true } },
    { text: "과거 성공 사례 찾기 어려움", options: { bullet: true, breakLine: true } },
    { text: "고객사별 맞춤화에 시간 소요", options: { bullet: true } },
  ], { x: 0.85, y: 1.85, w: 3.6, h: 2.2, fontSize: 13, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0, paraSpaceAfter: 6 });

  // After
  s6.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.15, w: 4.2, h: 3.2, fill: { color: CARD_BG }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.15, w: 0.08, h: 3.2, fill: { color: ACCENT_GREEN } });
  s6.addText("After", { x: 5.65, y: 1.3, w: 1.5, h: 0.4, fontSize: 18, fontFace: "Malgun Gothic", color: ACCENT_GREEN, bold: true, margin: 0 });
  s6.addText([
    { text: "고객사 정보(산업, 규모, 니즈) 입력", options: { bullet: true, breakLine: true } },
    { text: "과거 유사 사례 자동 검색", options: { bullet: true, breakLine: true } },
    { text: "맞춤 제안서 초안 자동 생성", options: { bullet: true, breakLine: true } },
    { text: "본인이 검토 · 다듬기만 수행", options: { bullet: true } },
  ], { x: 5.65, y: 1.85, w: 3.6, h: 2.2, fontSize: 13, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0, paraSpaceAfter: 6 });

  s6.addImage({ data: iconArrow, x: 4.82, y: 2.4, w: 0.35, h: 0.35 });

  // Reference
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 9, h: 0.7, fill: { color: SOFT_ORANGE_BG } });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 0.08, h: 0.7, fill: { color: CARROT_ORANGE } });
  s6.addText([
    { text: "참고 사례: ", options: { bold: true } },
    { text: "Thompson Advisory Group — 리서치 브리핑 준비 4~5시간 → 10~15분으로 단축" },
  ], { x: 0.85, y: 4.6, w: 8.4, h: 0.7, fontSize: 12, fontFace: "Malgun Gothic", color: CARROT_DARK, valign: "middle", margin: 0 });

  // ===== SLIDE 7: SCENARIO 2 - TRAINING MATERIALS =====
  let s7 = pres.addSlide();
  s7.background = { color: WHITE };
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: CARROT_ORANGE } });
  s7.addImage({ data: iconGradWhite, x: 0.5, y: 0.17, w: 0.48, h: 0.48 });
  s7.addText("적용 시나리오 ② 교육자료 자동화", {
    x: 1.1, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true,
    align: "left", valign: "middle", margin: 0
  });

  s7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 4.2, h: 3.2, fill: { color: CARD_BG }, shadow: makeShadow() });
  s7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 0.08, h: 3.2, fill: { color: ACCENT_RED } });
  s7.addText("Before", { x: 0.85, y: 1.3, w: 1.5, h: 0.4, fontSize: 18, fontFace: "Malgun Gothic", color: ACCENT_RED, bold: true, margin: 0 });
  s7.addText([
    { text: "고객사마다 교육자료 새로 제작", options: { bullet: true, breakLine: true } },
    { text: "산업별 맞춤이 시간 소요", options: { bullet: true, breakLine: true } },
    { text: "AI담당자 1명이 병목", options: { bullet: true, breakLine: true } },
    { text: "품질 편차 발생", options: { bullet: true } },
  ], { x: 0.85, y: 1.85, w: 3.6, h: 2.2, fontSize: 13, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0, paraSpaceAfter: 6 });

  s7.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.15, w: 4.2, h: 3.2, fill: { color: CARD_BG }, shadow: makeShadow() });
  s7.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.15, w: 0.08, h: 3.2, fill: { color: ACCENT_GREEN } });
  s7.addText("After", { x: 5.65, y: 1.3, w: 1.5, h: 0.4, fontSize: 18, fontFace: "Malgun Gothic", color: ACCENT_GREEN, bold: true, margin: 0 });
  s7.addText([
    { text: "산업별 레퍼런스 + 과거 자료 기반", options: { bullet: true, breakLine: true } },
    { text: "맞춤 교육 PPT 초안 자동 생성", options: { bullet: true, breakLine: true } },
    { text: "핸드아웃/실습자료도 자동화", options: { bullet: true, breakLine: true } },
    { text: "누구나 고품질 자료 제작 가능", options: { bullet: true } },
  ], { x: 5.65, y: 1.85, w: 3.6, h: 2.2, fontSize: 13, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0, paraSpaceAfter: 6 });

  s7.addImage({ data: iconArrow, x: 4.82, y: 2.4, w: 0.35, h: 0.35 });

  s7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 9, h: 0.7, fill: { color: SOFT_ORANGE_BG } });
  s7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 0.08, h: 0.7, fill: { color: CARROT_ORANGE } });
  s7.addText([
    { text: "참고 사례: ", options: { bold: true } },
    { text: "NIIT — SME 참여 시간 60% 절감  |  팀스파르타 — 콘텐츠 제작량 4배 증가" },
  ], { x: 0.85, y: 4.6, w: 8.4, h: 0.7, fontSize: 12, fontFace: "Malgun Gothic", color: CARROT_DARK, valign: "middle", margin: 0 });

  // ===== SLIDE 8: EXPECTED EFFECTS =====
  let s8 = pres.addSlide();
  s8.background = { color: WHITE };
  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: CARROT_ORANGE } });
  s8.addImage({ data: iconChartWhite, x: 0.5, y: 0.17, w: 0.48, h: 0.48 });
  s8.addText("기대 효과", {
    x: 1.1, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true,
    align: "left", valign: "middle", margin: 0
  });

  const metrics = [
    { stat: "60%", label: "산출물 제작 시간 단축", sub: "NIIT 사례 근거", icon: iconChartGreen, x: 0.5, y: 1.15 },
    { stat: "2~3배", label: "동시 수행 프로젝트 증가", sub: "자동화로 처리량 확대", icon: iconExpandBlue, x: 5.3, y: 1.15 },
    { stat: "Fast", label: "신규 입사자 온보딩 단축", sub: "소스 자료 즉시 활용 가능", icon: iconUsersOrange, x: 0.5, y: 3.2 },
    { stat: "1 → All", label: "AI 전문지식 의존도 분산", sub: "1명 의존 → 팀 전체로 확산", icon: iconRocketOrange, x: 5.3, y: 3.2 },
  ];
  for (const m of metrics) {
    s8.addShape(pres.shapes.RECTANGLE, { x: m.x, y: m.y, w: 4.2, h: 1.8, fill: { color: CARD_BG }, shadow: makeShadow() });
    s8.addShape(pres.shapes.RECTANGLE, { x: m.x, y: m.y, w: 0.08, h: 1.8, fill: { color: CARROT_ORANGE } });
    s8.addImage({ data: m.icon, x: m.x + 0.3, y: m.y + 0.35, w: 0.45, h: 0.45 });
    s8.addText(m.stat, { x: m.x + 0.9, y: m.y + 0.15, w: 3.0, h: 0.65, fontSize: 34, fontFace: "Malgun Gothic", color: CARROT_ORANGE, bold: true, margin: 0 });
    s8.addText(m.label, { x: m.x + 0.9, y: m.y + 0.8, w: 3.0, h: 0.35, fontSize: 14, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0 });
    s8.addText(m.sub, { x: m.x + 0.9, y: m.y + 1.15, w: 3.0, h: 0.3, fontSize: 11, fontFace: "Malgun Gothic", color: LIGHT_GRAY, margin: 0 });
  }

  // ===== SLIDE 9: EXPANSION =====
  let s9 = pres.addSlide();
  s9.background = { color: WHITE };
  s9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: CARROT_ORANGE } });
  s9.addImage({ data: iconRocketWhite, x: 0.5, y: 0.17, w: 0.48, h: 0.48 });
  s9.addText('확장 가능성: "써본 사람이 안다"', {
    x: 1.1, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true,
    align: "left", valign: "middle", margin: 0
  });

  // Core message
  s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 9, h: 0.9, fill: { color: SOFT_ORANGE_BG }, shadow: makeShadow() });
  s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 0.08, h: 0.9, fill: { color: CARROT_ORANGE } });
  s9.addText('"우리가 직접 쓰고 있습니다" = AX 컨설팅의 설득력이 달라진다', {
    x: 0.85, y: 1.15, w: 8.4, h: 0.9,
    fontSize: 17, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true,
    align: "center", valign: "middle", margin: 0
  });

  // Left card
  s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.35, w: 4.2, h: 2.8, fill: { color: CARD_BG }, shadow: makeShadow() });
  s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.35, w: 0.08, h: 2.8, fill: { color: CARROT_ORANGE } });
  s9.addText("AX 팀이 직접 활용하는 도구", { x: 0.85, y: 2.5, w: 3.6, h: 0.35, fontSize: 14, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0 });
  s9.addText([
    { text: "Slack — 팀 소통 · 프로젝트 관리", options: { bullet: true, breakLine: true } },
    { text: "Confluence — 지식 관리 · 소스 자료", options: { bullet: true, breakLine: true } },
    { text: "n8n / Dify — AI 워크플로우 자동화", options: { bullet: true, breakLine: true } },
    { text: "Claude — AI 업무 활용", options: { bullet: true } },
  ], { x: 0.85, y: 3.0, w: 3.6, h: 1.8, fontSize: 12, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0, paraSpaceAfter: 6 });

  // Right card
  s9.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 2.35, w: 4.2, h: 2.8, fill: { color: CARD_BG }, shadow: makeShadow() });
  s9.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 2.35, w: 0.08, h: 2.8, fill: { color: CARROT_DARK } });
  s9.addText("고객사가 쓰는 도구에도 적용", { x: 5.65, y: 2.5, w: 3.6, h: 0.35, fontSize: 14, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0 });
  s9.addText([
    { text: "Teams — LG, SK, 포스코", options: { bullet: true, breakLine: true } },
    { text: "SAP — 삼성, 현대, SK", options: { bullet: true, breakLine: true } },
    { text: "Jira — 카카오, 배민, 쿠팡", options: { bullet: true, breakLine: true } },
    { text: "Salesforce — LG화학, 롯데렌탈", options: { bullet: true } },
  ], { x: 5.65, y: 3.0, w: 3.6, h: 1.8, fontSize: 12, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0, paraSpaceAfter: 6 });

  // ===== SLIDE 10: Q&A =====
  let s10 = pres.addSlide();
  s10.background = { color: WHITE };
  s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: CARROT_ORANGE } });
  s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: CARROT_ORANGE } });

  s10.addText("Q & A", {
    x: 0, y: 1.5, w: 10, h: 1.2,
    fontSize: 52, fontFace: "Malgun Gothic", color: CARROT_ORANGE, bold: true,
    align: "center", valign: "middle", margin: 0
  });
  s10.addText("감사합니다", {
    x: 0, y: 2.8, w: 10, h: 0.8,
    fontSize: 24, fontFace: "Malgun Gothic", color: DARK_TEXT,
    align: "center", valign: "middle", margin: 0
  });

  s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.725, fill: { color: CARROT_ORANGE } });
  s10.addText("정찬인  |  캐럿글로벌 AX사업부 지원", {
    x: 0, y: 4.9, w: 10, h: 0.725,
    fontSize: 14, fontFace: "Malgun Gothic", color: WHITE,
    align: "center", valign: "middle", margin: 0
  });

  // Write file
  await pres.writeFile({ fileName: "C:/Users/ktrau/OneDrive/바탕 화면/interview/carrotglobal/ax_agent_proposal.pptx" });
  console.log("PPTX generated successfully!");
}

main().catch(err => { console.error(err); process.exit(1); });
