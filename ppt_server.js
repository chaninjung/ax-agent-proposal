const express = require("express");
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const app = express();
app.use(express.json());

const FOLDER_ID = "17sA0IlyIc_N2nJ4mFJjSO1BhkzwRmejB";
const GCLOUD = "C:/Users/ktrau/google-cloud-sdk/bin/gcloud.cmd";

function getAccessToken() {
  const raw = execSync(`${GCLOUD} auth print-access-token --scopes=https://www.googleapis.com/auth/drive`, { encoding: "utf-8" });
  return raw.replace("Python ", "").trim();
}

async function uploadToDrive(filePath, fileName) {
  const token = getAccessToken();
  const https = require("https");

  // Step 1: Create file metadata
  const metadata = JSON.stringify({ name: fileName, parents: [FOLDER_ID] });
  const fileData = fs.readFileSync(filePath);

  const boundary = "----FormBoundary" + Date.now();
  const metadataBuffer = Buffer.from(metadata, "utf-8");
  const body = Buffer.concat([
    Buffer.from(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`, "ascii"),
    metadataBuffer,
    Buffer.from(`\r\n--${boundary}\r\nContent-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation\r\n\r\n`, "ascii"),
    fileData,
    Buffer.from(`\r\n--${boundary}--`, "ascii"),
  ]);

  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: "www.googleapis.com",
      path: "/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink",
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": `multipart/related; boundary=${boundary}`,
        "Content-Length": body.length,
      },
    }, (res) => {
      let data = "";
      res.on("data", (chunk) => data += chunk);
      res.on("end", () => {
        console.log("Drive API response:", data);
        try {
          const result = JSON.parse(data);
          if (result.error) {
            reject(new Error(JSON.stringify(result.error)));
          } else {
            resolve({
              fileId: result.id,
              viewLink: result.webViewLink || `https://drive.google.com/file/d/${result.id}/view`,
            });
          }
        } catch (e) {
          reject(new Error(data));
        }
      });
    });
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}

// PPT Generation
async function generateProposalPPT(data) {
  const {
    client = "고객사",
    industry = "산업",
    topic = "AI 교육",
    target = "실무자",
    duration = "3시간",
    date = "TBD",
    ciColor = "E30613",
    content = "",
  } = data;

  const MAIN_COLOR = ciColor.replace("#", "");
  const DARK_TEXT = "2D2D2D";
  const MEDIUM_TEXT = "555555";
  const LIGHT_GRAY = "999999";
  const WHITE = "FFFFFF";
  const CARD_BG = "F7F7F7";

  const makeShadow = () => ({ type: "outer", blur: 4, offset: 1, angle: 135, color: "000000", opacity: 0.08 });

  let pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "캐럿글로벌 AX사업부";
  pres.title = `${client} AI 교육 제안서`;

  // === SLIDE 1: TITLE ===
  let s1 = pres.addSlide();
  s1.background = { color: WHITE };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: MAIN_COLOR } });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: MAIN_COLOR } });

  s1.addText(`${client}`, {
    x: 0.8, y: 1.4, w: 8.4, h: 0.6,
    fontSize: 18, fontFace: "Malgun Gothic", color: MAIN_COLOR, bold: true, align: "left", margin: 0,
  });
  s1.addText(`${topic} 교육 제안서`, {
    x: 0.8, y: 2.0, w: 8.4, h: 1.2,
    fontSize: 36, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, align: "left", margin: 0,
  });
  s1.addText(`${target} 대상 | ${duration} | ${date}`, {
    x: 0.8, y: 3.3, w: 8.4, h: 0.5,
    fontSize: 14, fontFace: "Malgun Gothic", color: LIGHT_GRAY, align: "left", margin: 0,
  });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.725, fill: { color: MAIN_COLOR } });
  s1.addText("캐럿글로벌 AX사업부", {
    x: 0.8, y: 4.9, w: 4, h: 0.725,
    fontSize: 14, fontFace: "Malgun Gothic", color: WHITE, align: "left", valign: "middle", margin: 0,
  });

  // === SLIDE 2: BACKGROUND ===
  let s2 = pres.addSlide();
  s2.background = { color: WHITE };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: MAIN_COLOR } });
  s2.addText("제안 배경", {
    x: 0.8, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true, align: "left", valign: "middle", margin: 0,
  });

  s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 3.8, fill: { color: CARD_BG }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 0.08, h: 3.8, fill: { color: MAIN_COLOR } });

  const bgText = `${client}은(는) ${industry} 분야의 대표 기업으로, AI 기술의 빠른 발전에 따라 전사적 AI 역량 강화가 필요한 시점입니다.\n\n특히 ${target} 대상의 ${topic} 교육을 통해:\n\n• AI 기술 트렌드와 비즈니스 영향력에 대한 이해\n• ${industry} 산업에서의 AI 활용 사례 학습\n• AI 기반 의사결정 역량 강화\n• 조직 내 AI 문화 확산을 위한 리더십 확보\n\n위 목표를 달성하고자 본 교육을 제안드립니다.`;
  s2.addText(bgText, {
    x: 0.85, y: 1.4, w: 8.4, h: 3.4,
    fontSize: 13, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0, lineSpacingMultiple: 1.3,
  });

  // === SLIDE 3: OBJECTIVES ===
  let s3 = pres.addSlide();
  s3.background = { color: WHITE };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: MAIN_COLOR } });
  s3.addText("교육 목표", {
    x: 0.8, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true, align: "left", valign: "middle", margin: 0,
  });

  const objectives = [
    { num: "01", title: "AI 이해도 향상", desc: `${target}의 AI 기본 개념 및\n최신 트렌드 이해` },
    { num: "02", title: "실무 적용 역량", desc: `${industry} 산업 특화\nAI 활용 사례 학습` },
    { num: "03", title: "조직 변화 리드", desc: `AI 기반 업무 혁신을\n주도할 리더십 확보` },
  ];

  objectives.forEach((obj, i) => {
    const x = 0.5 + i * 3.15;
    s3.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.85, h: 3.5, fill: { color: CARD_BG }, shadow: makeShadow() });
    s3.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 0.08, h: 3.5, fill: { color: MAIN_COLOR } });
    s3.addText(obj.num, { x: x + 0.25, y: 1.4, w: 0.8, h: 0.5, fontSize: 30, fontFace: "Malgun Gothic", color: MAIN_COLOR, bold: true, margin: 0 });
    s3.addText(obj.title, { x: x + 0.25, y: 2.0, w: 2.35, h: 0.4, fontSize: 16, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0 });
    s3.addShape(pres.shapes.RECTANGLE, { x: x + 0.25, y: 2.45, w: 1.0, h: 0.03, fill: { color: MAIN_COLOR } });
    s3.addText(obj.desc, { x: x + 0.25, y: 2.6, w: 2.35, h: 1.5, fontSize: 12, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, margin: 0 });
  });

  // === SLIDE 4: CURRICULUM ===
  let s4 = pres.addSlide();
  s4.background = { color: WHITE };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: MAIN_COLOR } });
  s4.addText("커리큘럼 구성", {
    x: 0.8, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true, align: "left", valign: "middle", margin: 0,
  });

  // Table
  const tableHeader = [
    { text: "시간", options: { fill: { color: MAIN_COLOR }, color: WHITE, bold: true, fontSize: 12, fontFace: "Malgun Gothic", align: "center", valign: "middle" } },
    { text: "주제", options: { fill: { color: MAIN_COLOR }, color: WHITE, bold: true, fontSize: 12, fontFace: "Malgun Gothic", align: "center", valign: "middle" } },
    { text: "내용", options: { fill: { color: MAIN_COLOR }, color: WHITE, bold: true, fontSize: 12, fontFace: "Malgun Gothic", align: "center", valign: "middle" } },
  ];

  const cellOpts = { fontSize: 11, fontFace: "Malgun Gothic", color: MEDIUM_TEXT, valign: "middle" };
  const tableRows = [
    ["30분", "AI 트렌드 브리핑", "생성형 AI 현황, 산업별 영향 분석"],
    ["60분", `${industry} AI 활용 사례`, `${industry} 분야 AI 도입 사례 및 성과`],
    ["30분", "AI 도구 체험", "ChatGPT, Claude 등 주요 AI 도구 실습"],
    ["30분", "AI 전략 토론", `${client} AI 도입 방향성 토론`],
    ["30분", "Q&A 및 마무리", "질의응답, 핵심 메시지 정리"],
  ];

  const tableData = [tableHeader];
  tableRows.forEach((row) => {
    tableData.push(row.map((text) => ({ text, options: { ...cellOpts } })));
  });

  s4.addTable(tableData, {
    x: 0.5, y: 1.2, w: 9, h: 3.5,
    border: { pt: 0.5, color: "E0E0E0" },
    colW: [1.5, 2.5, 5],
    rowH: [0.45, 0.55, 0.55, 0.55, 0.55, 0.55],
    autoPage: false,
  });

  // === SLIDE 5: EXPECTED EFFECTS ===
  let s5 = pres.addSlide();
  s5.background = { color: WHITE };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: MAIN_COLOR } });
  s5.addText("기대 효과", {
    x: 0.8, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true, align: "left", valign: "middle", margin: 0,
  });

  const effects = [
    { stat: "AI 이해", label: "전사 AI 리터러시 향상", sub: `${target} AI 활용 역량 확보` },
    { stat: "실무 적용", label: "업무 효율화 기반 마련", sub: `${industry} 특화 AI 적용 인사이트` },
    { stat: "문화 확산", label: "AI 중심 조직문화 조성", sub: "자발적 AI 활용 분위기 형성" },
    { stat: "경쟁력", label: "디지털 경쟁력 강화", sub: "AI 기반 비즈니스 혁신 가속화" },
  ];

  effects.forEach((e, i) => {
    const x = i < 2 ? 0.5 : 5.3;
    const y = i % 2 === 0 ? 1.15 : 3.2;
    s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.2, h: 1.7, fill: { color: CARD_BG }, shadow: makeShadow() });
    s5.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.08, h: 1.7, fill: { color: MAIN_COLOR } });
    s5.addText(e.stat, { x: x + 0.3, y: y + 0.15, w: 3.6, h: 0.55, fontSize: 24, fontFace: "Malgun Gothic", color: MAIN_COLOR, bold: true, margin: 0 });
    s5.addText(e.label, { x: x + 0.3, y: y + 0.7, w: 3.6, h: 0.35, fontSize: 14, fontFace: "Malgun Gothic", color: DARK_TEXT, bold: true, margin: 0 });
    s5.addText(e.sub, { x: x + 0.3, y: y + 1.05, w: 3.6, h: 0.3, fontSize: 11, fontFace: "Malgun Gothic", color: LIGHT_GRAY, margin: 0 });
  });

  // === SLIDE 6: ABOUT CARROT ===
  let s6 = pres.addSlide();
  s6.background = { color: WHITE };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: MAIN_COLOR } });
  s6.addText("캐럿글로벌 소개", {
    x: 0.8, y: 0, w: 8, h: 0.85,
    fontSize: 22, fontFace: "Malgun Gothic", color: WHITE, bold: true, align: "left", valign: "middle", margin: 0,
  });

  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 3.8, fill: { color: CARD_BG }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 0.08, h: 3.8, fill: { color: MAIN_COLOR } });
  s6.addText([
    { text: "캐럿글로벌 AX사업부\n", options: { fontSize: 18, bold: true, color: DARK_TEXT, breakLine: true } },
    { text: '"AI 도입의 본질은 기술이 아닌 사람과 업무의 재설계"\n\n', options: { fontSize: 13, italic: true, color: MAIN_COLOR, breakLine: true } },
    { text: "주요 수행 실적\n", options: { fontSize: 14, bold: true, color: DARK_TEXT, breakLine: true } },
    { text: "• 2026년 고용노동부 '중소기업 AI훈련확산센터' 수도권 수행기관 선정\n", options: { fontSize: 12, color: MEDIUM_TEXT, breakLine: true } },
    { text: "• GS리테일 전사 AI 교육 (100건 프로젝트)\n", options: { fontSize: 12, color: MEDIUM_TEXT, breakLine: true } },
    { text: "• 대상그룹 전사 AI 교육\n", options: { fontSize: 12, color: MEDIUM_TEXT, breakLine: true } },
    { text: "• 현대엔지니어링 AI 교육\n\n", options: { fontSize: 12, color: MEDIUM_TEXT, breakLine: true } },
    { text: "강점: ", options: { fontSize: 12, bold: true, color: DARK_TEXT } },
    { text: "실무 중심 교육 | 맞춤형 커리큘럼 | AX 전문 컨설팅 | 사후 적용 지원", options: { fontSize: 12, color: MEDIUM_TEXT } },
  ], { x: 0.85, y: 1.4, w: 8.4, h: 3.4, margin: 0 });

  // === SLIDE 7: THANK YOU ===
  let s7 = pres.addSlide();
  s7.background = { color: WHITE };
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: MAIN_COLOR } });
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: MAIN_COLOR } });
  s7.addText("감사합니다", {
    x: 0, y: 1.5, w: 10, h: 1.2,
    fontSize: 44, fontFace: "Malgun Gothic", color: MAIN_COLOR, bold: true, align: "center", valign: "middle", margin: 0,
  });
  s7.addText(`${client} ${topic} 교육 제안`, {
    x: 0, y: 2.8, w: 10, h: 0.8,
    fontSize: 18, fontFace: "Malgun Gothic", color: DARK_TEXT, align: "center", valign: "middle", margin: 0,
  });
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.9, w: 10, h: 0.725, fill: { color: MAIN_COLOR } });
  s7.addText("캐럿글로벌 AX사업부", {
    x: 0, y: 4.9, w: 10, h: 0.725,
    fontSize: 14, fontFace: "Malgun Gothic", color: WHITE, align: "center", valign: "middle", margin: 0,
  });

  // Save locally first - use safe ASCII filename for local, Korean for Drive
  const safeClient = client.replace(/[^a-zA-Z0-9_]/g, "") || "client";
  const safeTopic = topic.replace(/[^a-zA-Z0-9_]/g, "") || "topic";
  const timestamp = Date.now();
  const localFileName = `proposal_${safeClient}_${safeTopic}_${timestamp}.pptx`;
  const driveFileName = `${client}_${topic}_제안서_${timestamp}.pptx`;
  const fileName = localFileName;
  const filePath = path.join(__dirname, "output", fileName);
  if (!fs.existsSync(path.join(__dirname, "output"))) {
    fs.mkdirSync(path.join(__dirname, "output"), { recursive: true });
  }
  await pres.writeFile({ fileName: filePath });

  // Upload to Google Drive
  try {
    const driveResult = await uploadToDrive(filePath, driveFileName);
    return { fileName, downloadLink: driveResult.viewLink, location: "Google Drive" };
  } catch (err) {
    console.error("Drive upload failed, serving locally:", err.message);
    return { fileName, downloadLink: `/files/${encodeURIComponent(fileName)}`, location: "local" };
  }
}

// Serve static files from output directory
app.use("/files", express.static(path.join(__dirname, "output")));

// API Endpoint
app.post("/generate-ppt", async (req, res) => {
  try {
    console.log("Received request:", JSON.stringify(req.body, null, 2));
    const result = await generateProposalPPT(req.body);
    res.json({
      success: true,
      message: `PPT 제안서가 생성되었습니다!`,
      fileName: result.fileName,
      downloadLink: result.downloadLink,
    });
  } catch (err) {
    console.error("Error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

app.get("/openapi.json", (req, res) => {
  const host = req.headers.host || "localhost:3456";
  const scheme = req.headers["x-forwarded-proto"] || "http";
  res.json({
    "openapi": "3.0.0",
    "info": {
      "title": "PPT Generator API",
      "version": "1.0.0",
      "description": "AI 교육 제안서 PPT를 자동 생성하는 API"
    },
    "servers": [{ "url": `${scheme}://${host}` }],
    "paths": {
      "/generate-ppt": {
        "post": {
          "operationId": "generatePPT",
          "summary": "교육 제안서 PPT 생성",
          "description": "고객사 정보를 입력받아 맞춤형 교육 제안서 PPT를 생성합니다",
          "requestBody": {
            "required": true,
            "content": {
              "application/json": {
                "schema": {
                  "type": "object",
                  "required": ["client", "industry", "topic", "target"],
                  "properties": {
                    "client": { "type": "string", "description": "고객사명 (예: 대상_청정원)" },
                    "industry": { "type": "string", "description": "산업 분야 (예: 식품유통, 제조, 금융)" },
                    "topic": { "type": "string", "description": "교육 주제 (예: AI리터러시, AI업무자동화)" },
                    "target": { "type": "string", "description": "교육 대상 (예: 임원, 실무자, 전직원)" },
                    "duration": { "type": "string", "description": "교육 시간 (예: 3시간)" },
                    "date": { "type": "string", "description": "교육 일정 (예: 2025-05-06)" },
                    "ciColor": { "type": "string", "description": "고객사 CI 색상 hex코드 (예: E30613)" }
                  }
                }
              }
            }
          },
          "responses": {
            "200": {
              "description": "PPT 생성 성공",
              "content": {
                "application/json": {
                  "schema": {
                    "type": "object",
                    "properties": {
                      "success": { "type": "boolean" },
                      "message": { "type": "string" },
                      "fileName": { "type": "string" },
                      "downloadLink": { "type": "string" }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  });
});

const PORT = 3456;
app.listen(PORT, () => {
  console.log(`PPT Generation Server running on http://localhost:${PORT}`);
  console.log(`Health check: http://localhost:${PORT}/health`);
});
