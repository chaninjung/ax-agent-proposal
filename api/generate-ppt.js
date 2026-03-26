const pptxgen = require("pptxgenjs");

module.exports = async (req, res) => {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { client, industry, topic, target, duration, date, ciColor } = req.body;
    const primaryColor = ciColor || "FF6B2C";
    const darkText = "2D2D2D";
    const mediumText = "555555";
    const white = "FFFFFF";
    const lightGray = "F5F5F5";

    let pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";

    // Slide 1: Title
    let s1 = pres.addSlide();
    s1.background = { color: white };
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: primaryColor } });
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: primaryColor } });
    s1.addText("AI 교육 제안서", {
      x: 0.6, y: 1.5, w: 8.5, h: 0.5,
      fontSize: 16, fontFace: "Malgun Gothic", color: primaryColor, bold: true, margin: 0
    });
    s1.addText(`${client || "고객사"}\n${topic || "AI 교육"}`, {
      x: 0.6, y: 2.1, w: 8.5, h: 1.4,
      fontSize: 34, fontFace: "Malgun Gothic", color: darkText, bold: true, margin: 0
    });
    s1.addText(`${target || ""} 대상 | ${duration || ""} | ${date || ""}`, {
      x: 0.6, y: 3.6, w: 8.5, h: 0.4,
      fontSize: 13, fontFace: "Malgun Gothic", color: mediumText, margin: 0
    });
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.725, fill: { color: primaryColor } });
    s1.addText("캐럿글로벌 AX사업부", {
      x: 0.6, y: 4.9, w: 9, h: 0.725,
      fontSize: 14, fontFace: "Malgun Gothic", color: white, valign: "middle", margin: 0
    });

    // Slide 2: Background
    let s2 = pres.addSlide();
    s2.background = { color: white };
    s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: primaryColor } });
    s2.addText("제안 배경", {
      x: 0.6, y: 0, w: 9, h: 0.8,
      fontSize: 22, fontFace: "Malgun Gothic", color: white, bold: true, valign: "middle", margin: 0
    });
    s2.addText([
      { text: `${industry || "해당 산업"}의 AI 전환이 가속화되고 있습니다`, options: { bold: true, breakLine: true, fontSize: 16 } },
      { text: "", options: { breakLine: true, fontSize: 8 } },
      { text: `• ${client || "고객사"}는 ${industry || "산업"} 분야의 선도 기업으로, AI 기술의 전략적 활용이 경쟁력의 핵심이 되고 있습니다.`, options: { breakLine: true } },
      { text: `• ${target || "교육 대상"}의 AI 리터러시 확보는 조직 전체의 디지털 전환을 이끄는 첫 걸음입니다.`, options: { breakLine: true } },
      { text: `• 본 교육은 ${target || "교육 대상"}이 AI의 본질을 이해하고, 업무에 적용할 수 있는 실질적 역량을 갖추도록 설계되었습니다.`, options: { breakLine: true } },
    ], {
      x: 0.6, y: 1.2, w: 8.8, h: 3.5,
      fontSize: 14, fontFace: "Malgun Gothic", color: darkText, margin: 0, lineSpacingMultiple: 1.5
    });

    // Slide 3: Objectives
    let s3 = pres.addSlide();
    s3.background = { color: white };
    s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: primaryColor } });
    s3.addText("교육 목표", {
      x: 0.6, y: 0, w: 9, h: 0.8,
      fontSize: 22, fontFace: "Malgun Gothic", color: white, bold: true, valign: "middle", margin: 0
    });
    const objectives = [
      { title: "AI 개념 이해", desc: "생성형 AI의 원리와 한계를 정확히 이해하고 판단력 확보" },
      { title: "업무 적용력", desc: `${industry || "산업"} 분야에서 AI가 실제로 활용되는 사례 학습` },
      { title: "전략적 시야", desc: "AI 기반 의사결정과 조직 변화를 이끄는 리더십 역량 확보" },
    ];
    objectives.forEach((obj, i) => {
      const y = 1.1 + i * 1.4;
      s3.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 8.8, h: 1.15, fill: { color: lightGray } });
      s3.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.08, h: 1.15, fill: { color: primaryColor } });
      s3.addText(`0${i + 1}`, { x: 0.9, y: y + 0.1, w: 0.6, h: 0.4, fontSize: 22, fontFace: "Malgun Gothic", color: primaryColor, bold: true, margin: 0 });
      s3.addText(obj.title, { x: 1.6, y: y + 0.1, w: 7.5, h: 0.4, fontSize: 16, fontFace: "Malgun Gothic", color: darkText, bold: true, margin: 0 });
      s3.addText(obj.desc, { x: 1.6, y: y + 0.55, w: 7.5, h: 0.4, fontSize: 13, fontFace: "Malgun Gothic", color: mediumText, margin: 0 });
    });

    // Slide 4: Curriculum
    let s4 = pres.addSlide();
    s4.background = { color: white };
    s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: primaryColor } });
    s4.addText("커리큘럼 구성", {
      x: 0.6, y: 0, w: 9, h: 0.8,
      fontSize: 22, fontFace: "Malgun Gothic", color: white, bold: true, valign: "middle", margin: 0
    });
    const rows = [
      [
        { text: "시간", options: { bold: true, color: "FFFFFF", fill: { color: primaryColor } } },
        { text: "주제", options: { bold: true, color: "FFFFFF", fill: { color: primaryColor } } },
        { text: "내용", options: { bold: true, color: "FFFFFF", fill: { color: primaryColor } } },
      ],
      ["30분", "AI 패러다임의 변화", `생성형 AI의 등장과 ${industry || "산업"} 영향`],
      ["40분", "AI 핵심 기술 이해", "LLM, 프롬프트 엔지니어링, RAG 개념"],
      ["20분", "휴식", ""],
      ["40분", `${industry || "산업"} AI 활용 사례`, `${client || "고객사"} 맞춤 AI 적용 시나리오`],
      ["30분", "AI 시대의 리더십", "조직 AI 전환 전략과 의사결정 프레임워크"],
      ["20분", "Q&A 및 마무리", "핵심 정리 및 후속 액션 플랜"],
    ];
    s4.addTable(rows, {
      x: 0.6, y: 1.1, w: 8.8,
      border: { pt: 0.5, color: "DDDDDD" },
      colW: [1.2, 2.8, 4.8],
      fontSize: 12, fontFace: "Malgun Gothic", color: darkText,
      rowH: [0.45, 0.5, 0.5, 0.4, 0.5, 0.5, 0.5],
    });

    // Slide 5: Expected Effects
    let s5 = pres.addSlide();
    s5.background = { color: white };
    s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: primaryColor } });
    s5.addText("기대 효과", {
      x: 0.6, y: 0, w: 9, h: 0.8,
      fontSize: 22, fontFace: "Malgun Gothic", color: white, bold: true, valign: "middle", margin: 0
    });
    const effects = [
      { stat: "AI 이해도", value: "향상", desc: "임원진의 AI 기술 이해 및 판단력 강화" },
      { stat: "의사결정", value: "가속", desc: "AI 기반 데이터 드리븐 의사결정 문화 확산" },
      { stat: "조직 변화", value: "선도", desc: "AI 전환을 이끄는 리더십 역량 확보" },
      { stat: "경쟁력", value: "강화", desc: `${industry || "산업"} AI 선도 기업으로의 포지셔닝` },
    ];
    effects.forEach((eff, i) => {
      const x = 0.4 + (i % 2) * 4.8;
      const y = 1.1 + Math.floor(i / 2) * 2.0;
      s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.4, h: 1.7, fill: { color: lightGray } });
      s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.08, h: 1.7, fill: { color: primaryColor } });
      s5.addText(eff.value, { x: x + 0.3, y: y + 0.15, w: 1.5, h: 0.5, fontSize: 24, fontFace: "Malgun Gothic", color: primaryColor, bold: true, margin: 0 });
      s5.addText(eff.stat, { x: x + 1.8, y: y + 0.15, w: 2.3, h: 0.5, fontSize: 16, fontFace: "Malgun Gothic", color: darkText, bold: true, margin: 0 });
      s5.addText(eff.desc, { x: x + 0.3, y: y + 0.8, w: 3.8, h: 0.6, fontSize: 12, fontFace: "Malgun Gothic", color: mediumText, margin: 0 });
    });

    // Slide 6: About & Closing
    let s6 = pres.addSlide();
    s6.background = { color: white };
    s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.8, fill: { color: primaryColor } });
    s6.addText("캐럿글로벌 소개", {
      x: 0.6, y: 0, w: 9, h: 0.8,
      fontSize: 22, fontFace: "Malgun Gothic", color: white, bold: true, valign: "middle", margin: 0
    });
    s6.addText([
      { text: "캐럿글로벌은 기업교육 및 AX(AI Transformation) 전문 기업입니다.", options: { bold: true, breakLine: true, fontSize: 15 } },
      { text: "", options: { breakLine: true, fontSize: 8 } },
      { text: '"AI 도입의 본질은 기술이 아닌 사람과 업무의 재설계입니다"', options: { italic: true, breakLine: true, color: primaryColor } },
      { text: "", options: { breakLine: true, fontSize: 8 } },
      { text: "• 2026 고용노동부 '중소기업 AI훈련확산센터' 수도권 수행기관 선정", options: { breakLine: true } },
      { text: "• GS리테일 전사 AI 교육 100건 수행", options: { breakLine: true } },
      { text: "• 대상그룹, 현대엔지니어링 등 다수 대기업 AX 교육 수행", options: { breakLine: true } },
      { text: "• AI Agent 실습 공개과정 운영 (Google Gems + Make 활용)", options: { breakLine: true } },
    ], {
      x: 0.6, y: 1.2, w: 8.8, h: 3.5,
      fontSize: 13, fontFace: "Malgun Gothic", color: darkText, margin: 0, lineSpacingMultiple: 1.5
    });

    // Slide 7: Thank you
    let s7 = pres.addSlide();
    s7.background = { color: white };
    s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: primaryColor } });
    s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: primaryColor } });
    s7.addText("감사합니다", {
      x: 0, y: 1.8, w: 10, h: 1.0,
      fontSize: 40, fontFace: "Malgun Gothic", color: primaryColor, bold: true, align: "center", margin: 0
    });
    s7.addText(`${client || "고객사"} ${topic || "AI 교육"} 제안`, {
      x: 0, y: 2.9, w: 10, h: 0.5,
      fontSize: 16, fontFace: "Malgun Gothic", color: mediumText, align: "center", margin: 0
    });
    s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.725, fill: { color: primaryColor } });
    s7.addText("캐럿글로벌 AX사업부 | carrotglobal.com", {
      x: 0, y: 4.9, w: 10, h: 0.725,
      fontSize: 13, fontFace: "Malgun Gothic", color: white, align: "center", valign: "middle", margin: 0
    });

    // Generate as base64
    const pptxBase64 = await pres.write({ outputType: "base64" });
    const fileName = `${(client || "proposal").replace(/\s/g, "_")}_${(topic || "AI").replace(/\s/g, "_")}_proposal.pptx`;

    res.status(200).json({
      success: true,
      message: "PPT 제안서가 생성되었습니다!",
      fileName,
      downloadUrl: `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${pptxBase64}`,
      note: "base64 데이터로 제공됩니다. 다운로드하려면 링크를 클릭하세요."
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
};
