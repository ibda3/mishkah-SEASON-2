import express from "express";
import cors from "cors";
import { createServer as createViteServer } from "vite";
import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";

const app = express();
const PORT = 3000;
const DATA_FILE = path.join(process.cwd(), "game_data.json");
const EXCEL_FILE = path.join(process.cwd(), "mishkah_live.xlsx");
const CSV_STATUS_FILE = path.join(process.cwd(), "vmix_status.csv");
const CSV_RESULTS_FILE = path.join(process.cwd(), "vmix_results.csv");
const CSV_GAMECONTROL_FILE = path.join(process.cwd(), "vmix_gamecontrol.csv");

app.use(cors());
app.use(express.json({ limit: "50mb" }));

export interface PlayerAnswerRecord {
  choice: "A" | "B" | "C" | "D" | "";
  at: number | null;
}

export const getPlayerChoice = (ans: any): string => {
  if (!ans) return "";
  if (typeof ans === "string") return ans;
  if (typeof ans === "object" && typeof ans.choice === "string") return ans.choice;
  return "";
};

export const getPlayerTime = (ans: any): number | null => {
  if (!ans || typeof ans !== "object") return null;
  return typeof ans.at === "number" ? ans.at : null;
};

// --- Initial State (Matching user's Excel sheets: Status, GameControl, Results, Questions) ---
let gameState = {
  status: {
    currentQuestionId: 1,
    totalQuestions: 0,
    isVotingOpen: false,
    playerAnswers: Array(8).fill(null).map(() => ({ choice: "" as "A"|"B"|"C"|"D"|"", at: null as number | null })),
    votingOpenedAt: 0,
    timerEnd: 0,
    timerDuration: 20,
    correctAnswer: "",
    gameStarted: false,
    activeSegment: 1,
    scores: Array(8).fill(0),
    grandTotals: Array(8).fill(0),
    awardedQuestions: [] as number[],
    selection: {
      blue: Array(8).fill(0),
      green: Array(8).fill(0)
    }
  },
  playerNames: [
    "سمية أمير",
    "أسماء دليل",
    "لينة بن ناجي",
    "بسمة حلاب",
    "أيوب مصري",
    "أسامة الوناس",
    "زيد كموش",
    "راشدي ياسين"
  ],
  questions: [] as any[],
  gameControl: {
    seg1: Array(8).fill(0),
    seg2: Array(8).fill(0),
    seg3: Array(8).fill(0),
    seg3Count: Array(8).fill(0),
    seg3Bonus: Array(8).fill(0),
    seg4: Array(8).fill(0),
    seg5: Array(8).fill(0),
    live: Array(8).fill(0)
  },
  view: "seg1"
};

// --- Load Data from File ---
if (fs.existsSync(DATA_FILE)) {
  try {
    const savedData = JSON.parse(fs.readFileSync(DATA_FILE, "utf-8"));
    const rawAnswers = savedData.status?.playerAnswers;
    const normalizedAnswers: PlayerAnswerRecord[] = Array.isArray(rawAnswers)
      ? rawAnswers.map((a: any) => {
          if (typeof a === "object" && a !== null) {
            return {
              choice: (a.choice || "").toUpperCase() as any,
              at: typeof a.at === "number" ? a.at : null
            };
          }
          return {
            choice: (a || "").toString().toUpperCase() as any,
            at: null
          };
        })
      : Array(8).fill(null).map(() => ({ choice: "", at: null }));

    gameState = {
      ...gameState,
      ...savedData,
      status: {
        ...gameState.status,
        ...(savedData.status || {}),
        playerAnswers: normalizedAnswers,
        votingOpenedAt: savedData.status?.votingOpenedAt || 0,
        awardedQuestions: Array.isArray(savedData.status?.awardedQuestions) ? savedData.status.awardedQuestions : []
      },
      playerNames: savedData.playerNames || gameState.playerNames,
      gameControl: { ...gameState.gameControl, ...(savedData.gameControl || {}) }
    };
    console.log("✅ Game data loaded from local storage");
  } catch (e) {
    console.error("❌ Error loading game data:", e);
  }
}

// --- Local vMix Excel & CSV Generator (Identical to User's Images) ---
const generateVmixFiles = () => {
  try {
    const wb = XLSX.utils.book_new();

    // 1. Status Sheet (Matching Image 2: Status with Answer Times Addendum)
    const currentQ = gameState.questions[gameState.status.currentQuestionId - 1];
    const correctLetter = currentQ ? currentQ.correct : (gameState.status.correctAnswer || "");
    const answeredIndices = gameState.status.playerAnswers
      .map((ans, idx) => (getPlayerChoice(ans) ? idx : null))
      .filter((idx) => idx !== null)
      .join(",");
    const timerSync = `${gameState.status.currentQuestionId}|${answeredIndices || "0,1,2,3,4,5,6,7"}`;

    const choices = gameState.status.playerAnswers.map(ans => getPlayerChoice(ans));
    const times = gameState.status.playerAnswers.map(ans => {
      const at = getPlayerTime(ans);
      if (at && gameState.status.votingOpenedAt) {
        const elapsed = Math.max(0, (at - gameState.status.votingOpenedAt) / 1000);
        return `${elapsed.toFixed(2)}s`;
      }
      return "";
    });

    const statusAoa: any[][] = [];
    // Row 1: Headers (A1 to P1, followed by P1_time to P8_time backward compatible)
    statusAoa.push([
      "QuestionId", "VotingOpen", "ShowAns",
      "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8",
      "", "", "CorrectAnswer", "", timerSync,
      "P1_time", "P2_time", "P3_time", "P4_time", "P5_time", "P6_time", "P7_time", "P8_time"
    ]);

    // Row 2: Values (A2 to N2, followed by P1_time to P8_time)
    statusAoa.push([
      gameState.status.currentQuestionId,
      gameState.status.isVotingOpen ? 1 : 0,
      0, // ShowAns
      ...choices,
      "", "", correctLetter, "",
      ...times
    ]);

    // Row 3: Correct/Answer Indicators (D3 to K3)
    statusAoa.push([
      "", "", "",
      ...choices.map(ans => (ans && correctLetter && ans.toUpperCase() === correctLetter.toUpperCase() ? 1 : 0))
    ]);

    // Row 4: Spacer
    statusAoa.push([]);

    // Row 5: Total Score Blue Row (D5 to K5)
    statusAoa.push([
      "", "", "",
      ...gameState.status.grandTotals
    ]);

    // Row 6: Win status (D6 to K6)
    statusAoa.push([
      "", "", "",
      ...gameState.status.grandTotals.map(() => "WIN")
    ]);

    const wsStatus = XLSX.utils.aoa_to_sheet(statusAoa);
    XLSX.utils.book_append_sheet(wb, wsStatus, "Status");

    // 2. GameControl Sheet (Matching Image 3: GameControl)
    const gcAoa: any[][] = [];
    // Row 1: Player headers
    gcAoa.push(["", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", ""]);
    // Rows 2 to 8: Part 1..4 & Total
    gcAoa.push(["", ...gameState.gameControl.seg1, "PART 1"]);
    gcAoa.push(["", ...gameState.gameControl.seg2, "PART 2"]);
    gcAoa.push(["", ...gameState.gameControl.seg3, "PART 3"]);
    gcAoa.push(["", ...gameState.gameControl.seg3Count, "...."]);
    gcAoa.push(["", ...gameState.gameControl.seg3Bonus, "PART 3 (+5)"]);
    gcAoa.push(["", ...gameState.gameControl.seg4, "PART 4"]);
    gcAoa.push(["", ...gameState.status.grandTotals, "TOTAL"]);

    // Rows 9 to 14: Part 5 & Live screen row
    gcAoa.push(["", ...gameState.gameControl.seg5, "PART 5"]);
    gcAoa.push(["", 0, 0, 0, 0, 0, 0, 0, 0, ""]);
    gcAoa.push(["", 0, 0, 0, 0, 0, 0, 0, 0, ""]);
    gcAoa.push([`LIVE SCREEN (${gameState.view})`, 0, 0, 0, 0, 0, 0, 0, 0, ""]);
    gcAoa.push(["", 0, 0, 0, 0, 0, 0, 0, 0, ""]);
    gcAoa.push(["", 0, 0, 0, 0, 0, 0, 0, 0, ""]);

    // Rows 15 to 20
    for (let i = 15; i <= 20; i++) {
      gcAoa.push(["", 0, 0, 0, 0, 0, 0, 0, 0, ""]);
    }
    // Row 21: Blue Selection Row ("take")
    gcAoa.push(["", ...gameState.status.selection.blue, "take"]);
    // Row 22: Green Selection Row ("action")
    gcAoa.push(["", ...gameState.status.selection.green, "action"]);

    const wsGC = XLSX.utils.aoa_to_sheet(gcAoa);
    XLSX.utils.book_append_sheet(wb, wsGC, "GameControl");

    // 3. Results Sheet (Matching Image 4: Results)
    const resAoa: any[][] = [];
    resAoa.push([
      "Person", "NAME", "Part 01", "Part 02", "Part 03", "Part 04", "Part 05", "TOTAL",
      "", "", "ActivePart", "", "FootagePreview"
    ]);

    for (let i = 0; i < 8; i++) {
      const rawName = gameState.playerNames?.[i];
      const pName = rawName !== undefined ? rawName : `المتسابق ${i + 1}`;
      const part3Total = (gameState.gameControl.seg3[i] || 0) + (gameState.gameControl.seg3Bonus[i] || 0);
      const activePartNum = gameState.status.activeSegment || 1;
      const previewImg = `Part 0${activePartNum}.png`;

      resAoa.push([
        `P${i + 1}`,
        pName,
        gameState.gameControl.seg1[i] || 0,
        gameState.gameControl.seg2[i] || 0,
        part3Total,
        gameState.gameControl.seg4[i] || 0,
        gameState.gameControl.seg5[i] || 0,
        gameState.status.grandTotals[i] || 0,
        "", "",
        activePartNum,
        "",
        previewImg
      ]);
    }
    const wsResults = XLSX.utils.aoa_to_sheet(resAoa);
    XLSX.utils.book_append_sheet(wb, wsResults, "Results");

    // 4. Questions Sheet (Matching Image 1: Questions)
    const qAoa: any[][] = [];
    qAoa.push(["Question", "A", "B", "C", "D", "Correct"]);
    gameState.questions.forEach((q: any) => {
      qAoa.push([q.text || "", q.a || "", q.b || "", q.c || "", q.d || "", q.correct || ""]);
    });
    const wsQuestions = XLSX.utils.aoa_to_sheet(qAoa);
    XLSX.utils.book_append_sheet(wb, wsQuestions, "Questions");

    // Write Multi-Sheet Excel File to disk
    XLSX.writeFile(wb, EXCEL_FILE);

    // Also write CSV files for direct single-table linking in vMix
    fs.writeFileSync(CSV_STATUS_FILE, XLSX.utils.sheet_to_csv(wsStatus), "utf-8");
    fs.writeFileSync(CSV_RESULTS_FILE, XLSX.utils.sheet_to_csv(wsResults), "utf-8");
    fs.writeFileSync(CSV_GAMECONTROL_FILE, XLSX.utils.sheet_to_csv(wsGC), "utf-8");

  } catch (e) {
    console.error("❌ Error generating local vMix Excel/CSV files:", e);
  }
};

// --- Save Data to File & Update vMix Excel/CSV ---
const saveGameData = () => {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(gameState, null, 2));
    generateVmixFiles();
  } catch (e) {
    console.error("❌ Error saving game data:", e);
  }
};

// Generate files initially if not present
generateVmixFiles();

// --- Helper Functions ---
const updateGrandTotals = () => {
  for (let i = 0; i < 8; i++) {
    gameState.status.grandTotals[i] =
      (gameState.gameControl.seg1[i] || 0) +
      (gameState.gameControl.seg2[i] || 0) +
      (gameState.gameControl.seg3[i] || 0) +
      (gameState.gameControl.seg3Bonus[i] || 0) +
      (gameState.gameControl.seg4[i] || 0) +
      (gameState.gameControl.seg5[i] || 0);
  }
};

const updateLiveRow = (viewMode: string) => {
  for (let i = 0; i < 8; i++) {
    let val = 0;
    if (viewMode === "seg1") val = gameState.gameControl.seg1[i] || 0;
    else if (viewMode === "seg2") val = gameState.gameControl.seg2[i] || 0;
    else if (viewMode === "seg3") val = (gameState.gameControl.seg3[i] || 0) + (gameState.gameControl.seg3Bonus[i] || 0);
    else if (viewMode === "seg4") val = gameState.gameControl.seg4[i] || 0;
    else if (viewMode === "seg5") val = gameState.gameControl.seg5[i] || 0;
    else if (viewMode === "total") val = gameState.status.grandTotals[i] || 0;
    gameState.gameControl.live[i] = val;
  }
};

const resetQuestionState = () => {
  gameState.status.isVotingOpen = true; // Auto-open voting on next/prev question
  gameState.status.playerAnswers = Array(8).fill(null).map(() => ({ choice: "" as const, at: null }));
  const duration = gameState.status.timerDuration || 20;
  const now = Date.now();
  gameState.status.votingOpenedAt = now;
  gameState.status.timerEnd = now + duration * 1000;
  const q = gameState.questions[gameState.status.currentQuestionId - 1];
  gameState.status.correctAnswer = q ? q.correct : "";
};

// --- API Routes ---

const setNoCacheHeaders = (res: any) => {
  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
};

// Lightweight endpoint for accurate client clock offset calculation (server-authoritative time)
app.get("/api/time", (req, res) => {
  setNoCacheHeaders(res);
  res.json({ success: true, serverTime: Date.now() });
});

// Get Status (Includes player names & local file paths & live vMix URLs)
app.get("/api/getStatus", (req, res) => {
  setNoCacheHeaders(res);
  const q = gameState.questions[gameState.status.currentQuestionId - 1] || null;
  const protocol = req.headers["x-forwarded-proto"] || req.protocol || "http";
  const host = req.get("host") || `localhost:${PORT}`;
  const baseUrl = `${protocol}://${host}`;

  res.json({
    success: true,
    serverTime: Date.now(),
    data: {
      status: {
        ...gameState.status,
        totalQuestions: gameState.questions.length
      },
      playerNames: gameState.playerNames,
      question: q,
      localFiles: {
        excel: EXCEL_FILE,
        csvStatus: CSV_STATUS_FILE,
        csvResults: CSV_RESULTS_FILE,
        csvGameControl: CSV_GAMECONTROL_FILE
      },
      vmixUrls: {
        resultsCsv: `${baseUrl}/api/vmix/results.csv`,
        statusCsv: `${baseUrl}/api/vmix/status.csv`,
        gamecontrolCsv: `${baseUrl}/api/vmix/gamecontrol.csv`,
        liveJson: `${baseUrl}/api/vmix/live.json`,
        excel: `${baseUrl}/mishkah_live.xlsx`
      }
    }
  });
});

// Download or link live Excel file for vMix
app.get(["/api/exportExcel", "/mishkah_live.xlsx", "/api/vmix/mishkah_live.xlsx"], (req, res) => {
  setNoCacheHeaders(res);
  if (fs.existsSync(EXCEL_FILE)) {
    res.download(EXCEL_FILE, "mishkah_live.xlsx");
  } else {
    generateVmixFiles();
    if (fs.existsSync(EXCEL_FILE)) {
      res.download(EXCEL_FILE, "mishkah_live.xlsx");
    } else {
      res.status(404).send("File not found");
    }
  }
});

// Download or view live CSV files with UTF-8 BOM for Arabic support in Excel & vMix
app.get("/api/vmix/status.csv", (req, res) => {
  setNoCacheHeaders(res);
  if (fs.existsSync(CSV_STATUS_FILE)) {
    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.sendFile(CSV_STATUS_FILE);
  } else {
    generateVmixFiles();
    res.sendFile(CSV_STATUS_FILE);
  }
});

app.get("/api/vmix/results.csv", (req, res) => {
  setNoCacheHeaders(res);
  if (fs.existsSync(CSV_RESULTS_FILE)) {
    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.sendFile(CSV_RESULTS_FILE);
  } else {
    generateVmixFiles();
    res.sendFile(CSV_RESULTS_FILE);
  }
});

app.get("/api/vmix/gamecontrol.csv", (req, res) => {
  setNoCacheHeaders(res);
  if (fs.existsSync(CSV_GAMECONTROL_FILE)) {
    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.sendFile(CSV_GAMECONTROL_FILE);
  } else {
    generateVmixFiles();
    res.sendFile(CSV_GAMECONTROL_FILE);
  }
});

// Live structured JSON data source for vMix
// Live structured JSON data source for vMix
app.get("/api/vmix/live.json", (req, res) => {
  setNoCacheHeaders(res);
  const q = gameState.questions[gameState.status.currentQuestionId - 1] || null;
  const contestants = [];
  for (let i = 0; i < 8; i++) {
    const rawName = gameState.playerNames?.[i] || "";
    const displayName = rawName.trim() || `المتسابق ${i + 1}`;
    const ansObj = gameState.status.playerAnswers[i];
    const choice = getPlayerChoice(ansObj);
    const at = getPlayerTime(ansObj);
    const elapsed = at && gameState.status.votingOpenedAt ? Math.max(0, (at - gameState.status.votingOpenedAt) / 1000) : null;
    contestants.push({
      id: i + 1,
      name: displayName,
      rawName: rawName,
      score: gameState.status.grandTotals[i] || 0,
      part1: gameState.gameControl.seg1[i] || 0,
      part2: gameState.gameControl.seg2[i] || 0,
      part3: (gameState.gameControl.seg3[i] || 0) + (gameState.gameControl.seg3Bonus[i] || 0),
      part3Count: gameState.gameControl.seg3Count[i] || 0,
      part3Bonus: gameState.gameControl.seg3Bonus[i] || 0,
      part4: gameState.gameControl.seg4[i] || 0,
      part5: gameState.gameControl.seg5[i] || 0,
      total: gameState.status.grandTotals[i] || 0,
      answer: choice,
      answerAt: at,
      elapsedTime: elapsed !== null ? `${elapsed.toFixed(2)}s` : null,
      isSelected: gameState.status.selection.blue[i] === 1 ? 1 : 0
    });
  }

  res.json({
    updatedAt: new Date().toISOString(),
    serverTime: Date.now(),
    currentQuestionId: gameState.status.currentQuestionId,
    totalQuestions: gameState.questions.length,
    activeSegment: gameState.status.activeSegment,
    isVotingOpen: gameState.status.isVotingOpen,
    votingOpenedAt: gameState.status.votingOpenedAt,
    timerEnd: gameState.status.timerEnd,
    timerDuration: gameState.status.timerDuration,
    correctAnswer: gameState.status.correctAnswer || (q ? q.correct : ""),
    questionText: q ? q.text : "",
    contestants
  });
});

// Submit Answer (Server is final authority: rejects any answer received after timerEnd)
app.get("/api/submitAnswer", (req, res) => {
  const { playerId, answer } = req.query;
  const pIdx = (Number(playerId) || 1) - 1;
  const ans = String(answer || "").trim().toUpperCase();
  const now = Date.now();

  if (gameState.status.gameStarted && gameState.status.isVotingOpen) {
    if (gameState.status.timerEnd > 0 && now > gameState.status.timerEnd) {
      return res.status(400).json({ success: false, error: "Time expired", rejected: true });
    }
    const currentAns = gameState.status.playerAnswers[pIdx];
    const existingChoice = getPlayerChoice(currentAns);
    if (!existingChoice) {
      gameState.status.playerAnswers[pIdx] = {
        choice: ans as any,
        at: now
      };
      saveGameData();
    }
  }
  res.json({ success: true });
});

// Toggle Voting with Countdown Timer support & start timestamp
app.get("/api/toggleVoting", (req, res) => {
  const openQuery = req.query.open;
  if (openQuery !== undefined) {
    gameState.status.isVotingOpen = openQuery === "true";
  } else {
    gameState.status.isVotingOpen = !gameState.status.isVotingOpen;
  }

  const now = Date.now();
  if (gameState.status.isVotingOpen) {
    const duration = Number(req.query.duration) || gameState.status.timerDuration || 20;
    gameState.status.timerDuration = duration;
    gameState.status.votingOpenedAt = now;
    gameState.status.timerEnd = now + duration * 1000;
  } else {
    gameState.status.timerEnd = 0;
  }
  saveGameData();
  res.json({
    success: true,
    isVotingOpen: gameState.status.isVotingOpen,
    votingOpenedAt: gameState.status.votingOpenedAt,
    timerEnd: gameState.status.timerEnd,
    timerDuration: gameState.status.timerDuration
  });
});

// Configure Timer Duration
app.get("/api/setTimerDuration", (req, res) => {
  const duration = Math.max(5, Math.min(180, Number(req.query.duration) || 20));
  gameState.status.timerDuration = duration;
  if (gameState.status.isVotingOpen) {
    gameState.status.timerEnd = Date.now() + duration * 1000;
  }
  saveGameData();
  res.json({ success: true, timerDuration: duration, timerEnd: gameState.status.timerEnd });
});

// Add extra seconds during live broadcast
app.get("/api/addTimerSeconds", (req, res) => {
  const extra = Number(req.query.seconds) || 5;
  if (gameState.status.isVotingOpen) {
    const now = Date.now();
    const currentEnd = Math.max(now, gameState.status.timerEnd || now);
    gameState.status.timerEnd = currentEnd + extra * 1000;
  }
  saveGameData();
  res.json({ success: true, timerEnd: gameState.status.timerEnd });
});

// Set Game Status
app.get("/api/setGameStatus", (req, res) => {
  const started = req.query.started === "true";
  gameState.status.gameStarted = started;
  if (!started) gameState.status.isVotingOpen = false;
  saveGameData();
  res.json({ success: true });
});

// Next/Prev Question
app.get("/api/nextQuestion", (req, res) => {
  gameState.status.currentQuestionId = Math.min(gameState.status.currentQuestionId + 1, gameState.questions.length || 1);
  resetQuestionState();
  saveGameData();
  res.json({ success: true });
});

app.get("/api/prevQuestion", (req, res) => {
  gameState.status.currentQuestionId = Math.max(1, gameState.status.currentQuestionId - 1);
  resetQuestionState();
  saveGameData();
  res.json({ success: true });
});

// Update Scores (Quiz Logic - Awarded ONLY ONCE per question)
app.get("/api/updateScores", (req, res) => {
  const qId = gameState.status.currentQuestionId;
  if (!Array.isArray(gameState.status.awardedQuestions)) {
    gameState.status.awardedQuestions = [];
  }

  // Prevent multiple score awards for the same question
  if (gameState.status.awardedQuestions.includes(qId)) {
    return res.json({
      success: false,
      alreadyAwarded: true,
      message: `تم احتساب درجات السؤال رقم (${qId}) مسبقاً! تحتسب الدرجات مرة واحدة فقط لكل سؤال لمنع تكرار النقاط.`
    });
  }

  const q = gameState.questions[qId - 1];
  if (q && q.correct) {
    let correctCount = 0;
    gameState.status.playerAnswers.forEach((ansObj, i) => {
      const choice = getPlayerChoice(ansObj);
      if (choice && choice.trim().toUpperCase() === q.correct.trim().toUpperCase()) {
        gameState.status.scores[i] = (gameState.status.scores[i] || 0) + 10;
        gameState.gameControl.seg1[i] = gameState.status.scores[i];
        correctCount++;
      }
    });

    gameState.status.awardedQuestions.push(qId);
    gameState.status.isVotingOpen = false;
    gameState.status.timerEnd = 0;

    updateGrandTotals();
    updateLiveRow(gameState.view);
    saveGameData();

    return res.json({
      success: true,
      message: `تم احتساب درجات السؤال (${qId}) بنجاح (+10 نقاط لـ ${correctCount} متسابق)`,
      awardedQuestions: gameState.status.awardedQuestions
    });
  }

  res.json({ success: false, message: "لم يتم العثور على الإجابة النموذجية لهذا السؤال" });
});

// Reset Score for a specific question (Undo feature)
app.get("/api/resetQuestionScore", (req, res) => {
  const qId = Number(req.query.questionId) || gameState.status.currentQuestionId;
  if (!Array.isArray(gameState.status.awardedQuestions)) {
    gameState.status.awardedQuestions = [];
  }

  if (!gameState.status.awardedQuestions.includes(qId)) {
    return res.json({ success: false, message: "درجات هذا السؤال لم تحتسب بعد" });
  }

  const q = gameState.questions[qId - 1];
  if (q && q.correct) {
    gameState.status.playerAnswers.forEach((ansObj, i) => {
      const choice = getPlayerChoice(ansObj);
      if (choice && choice.trim().toUpperCase() === q.correct.trim().toUpperCase()) {
        gameState.status.scores[i] = Math.max(0, (gameState.status.scores[i] || 0) - 10);
        gameState.gameControl.seg1[i] = gameState.status.scores[i];
      }
    });
  }

  gameState.status.awardedQuestions = gameState.status.awardedQuestions.filter(id => id !== qId);
  updateGrandTotals();
  updateLiveRow(gameState.view);
  saveGameData();

  res.json({
    success: true,
    message: `تم إلغاء احتساب درجات السؤال ${qId} بنجاح، ويمكن الآن إعادة احتسابه`,
    awardedQuestions: gameState.status.awardedQuestions
  });
});

// Segment Controls
app.get("/api/setActiveSegment", (req, res) => {
  gameState.status.activeSegment = Number(req.query.segment as string) || 1;
  saveGameData();
  res.json({ success: true });
});

// Segment Points Operations (Add & Subtract with Part Awareness)
app.get("/api/addSeg2", (req, res) => {
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  gameState.gameControl.seg2[pIdx] = (gameState.gameControl.seg2[pIdx] || 0) + 20;
  updateLiveRow("seg2");
  updateGrandTotals();
  saveGameData();
  res.json({ success: true });
});

app.get("/api/subSeg2", (req, res) => {
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  gameState.gameControl.seg2[pIdx] = Math.max(0, (gameState.gameControl.seg2[pIdx] || 0) - 20);
  updateLiveRow("seg2");
  updateGrandTotals();
  saveGameData();
  res.json({ success: true });
});

app.get("/api/addSeg3", (req, res) => {
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  gameState.gameControl.seg3Count[pIdx] = Math.min(7, (gameState.gameControl.seg3Count[pIdx] || 0) + 1);
  gameState.gameControl.seg3[pIdx] = gameState.gameControl.seg3Count[pIdx] * 10;
  if (gameState.gameControl.seg3Count[pIdx] === 7) {
    gameState.gameControl.seg3Bonus[pIdx] = 5;
  }
  updateLiveRow("seg3");
  updateGrandTotals();
  saveGameData();
  res.json({ success: true, count: gameState.gameControl.seg3Count[pIdx], bonus: gameState.gameControl.seg3Bonus[pIdx] });
});

app.get("/api/subSeg3", (req, res) => {
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  gameState.gameControl.seg3Count[pIdx] = Math.max(0, (gameState.gameControl.seg3Count[pIdx] || 0) - 1);
  gameState.gameControl.seg3[pIdx] = gameState.gameControl.seg3Count[pIdx] * 10;
  if (gameState.gameControl.seg3Count[pIdx] < 7) {
    gameState.gameControl.seg3Bonus[pIdx] = 0;
  }
  updateLiveRow("seg3");
  updateGrandTotals();
  saveGameData();
  res.json({ success: true, count: gameState.gameControl.seg3Count[pIdx], bonus: gameState.gameControl.seg3Bonus[pIdx] });
});

app.get("/api/addSeg4", (req, res) => {
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  gameState.gameControl.seg4[pIdx] = (gameState.gameControl.seg4[pIdx] || 0) + 25;
  updateLiveRow("seg4");
  updateGrandTotals();
  saveGameData();
  res.json({ success: true });
});

app.get("/api/subSeg4", (req, res) => {
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  gameState.gameControl.seg4[pIdx] = Math.max(0, (gameState.gameControl.seg4[pIdx] || 0) - 25);
  updateLiveRow("seg4");
  updateGrandTotals();
  saveGameData();
  res.json({ success: true });
});

app.get("/api/addSeg5", (req, res) => {
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  const pts = Number(req.query.points) || 0;
  gameState.gameControl.seg5[pIdx] = Math.max(0, (gameState.gameControl.seg5[pIdx] || 0) + pts);
  updateLiveRow("seg5");
  updateGrandTotals();
  saveGameData();
  res.json({ success: true, newScore: gameState.gameControl.seg5[pIdx] });
});

app.get("/api/toggleSelect", (req, res) => {
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  const nextVal = gameState.status.selection.blue[pIdx] === 1 ? 0 : 1;
  gameState.status.selection.blue[pIdx] = nextVal;
  gameState.status.selection.green[pIdx] = nextVal; // 1-click sync for vMix
  saveGameData();
  res.json({ success: true, isSelected: nextVal === 1 });
});

app.get("/api/commitSelection", (req, res) => {
  gameState.status.selection.green = [...gameState.status.selection.blue];
  saveGameData();
  res.json({ success: true });
});

app.get("/api/adjustPlayerScore", (req, res) => {
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  const delta = Number(req.query.delta) || 0;
  const seg = gameState.status.activeSegment || 1;
  if (seg === 1) gameState.gameControl.seg1[pIdx] = Math.max(0, (gameState.gameControl.seg1[pIdx] || 0) + delta);
  else if (seg === 2) gameState.gameControl.seg2[pIdx] = Math.max(0, (gameState.gameControl.seg2[pIdx] || 0) + delta);
  else if (seg === 3) gameState.gameControl.seg3[pIdx] = Math.max(0, (gameState.gameControl.seg3[pIdx] || 0) + delta);
  else if (seg === 4) gameState.gameControl.seg4[pIdx] = Math.max(0, (gameState.gameControl.seg4[pIdx] || 0) + delta);
  else if (seg === 5) gameState.gameControl.seg5[pIdx] = Math.max(0, (gameState.gameControl.seg5[pIdx] || 0) + delta);
  updateGrandTotals();
  updateLiveRow(gameState.view);
  saveGameData();
  res.json({ success: true });
});

app.get("/api/setView", (req, res) => {
  gameState.view = (req.query.view as string) || "seg1";
  updateLiveRow(gameState.view);
  saveGameData();
  res.json({ success: true });
});

app.get("/api/resetSegment", (req, res) => {
  const seg = req.query.segment as string;
  for (let i = 0; i < 8; i++) {
    if (seg === "1") {
      gameState.status.scores[i] = 0;
      gameState.gameControl.seg1[i] = 0;
    }
    if (seg === "2") gameState.gameControl.seg2[i] = 0;
    if (seg === "3") {
      gameState.gameControl.seg3[i] = 0;
      gameState.gameControl.seg3Count[i] = 0;
      gameState.gameControl.seg3Bonus[i] = 0;
    }
    if (seg === "4") gameState.gameControl.seg4[i] = 0;
    if (seg === "5") gameState.gameControl.seg5[i] = 0;
  }
  if (seg === "1") {
    gameState.status.awardedQuestions = [];
  }
  updateGrandTotals();
  updateLiveRow(gameState.view);
  saveGameData();
  res.json({ success: true });
});

// Update Contestant Names (Supports JSON body or query param)
app.all("/api/updatePlayerNames", (req, res) => {
  setNoCacheHeaders(res);
  let rawNames = req.body?.names;
  if (!rawNames && req.query?.names) {
    try {
      rawNames = JSON.parse(req.query.names as string);
    } catch {
      rawNames = null;
    }
  }

  if (Array.isArray(rawNames)) {
    const cleanedNames: string[] = [];
    for (let i = 0; i < 8; i++) {
      const val = rawNames[i];
      // Save exactly what the user typed (even empty string if cleared)
      cleanedNames[i] = typeof val === "string" ? val.trim() : "";
    }
    gameState.playerNames = cleanedNames;
    saveGameData();
    return res.json({ success: true, playerNames: gameState.playerNames });
  }

  res.status(400).json({ error: "Invalid player names (must be array)" });
});

// Update a single player name by index
app.get("/api/updateSinglePlayerName", (req, res) => {
  setNoCacheHeaders(res);
  const pIdx = (Number(req.query.playerId) || 1) - 1;
  const name = typeof req.query.name === "string" ? req.query.name.trim() : "";
  if (pIdx >= 0 && pIdx < 8) {
    gameState.playerNames[pIdx] = name;
    saveGameData();
    return res.json({ success: true, playerNames: gameState.playerNames });
  }
  res.status(400).json({ error: "Invalid player index" });
});

// Reset Contestant Names to Defaults
app.get("/api/resetPlayerNames", (req, res) => {
  setNoCacheHeaders(res);
  gameState.playerNames = [
    "سمية أمير",
    "أسماء دليل",
    "لينة بن ناجي",
    "بسمة حلاب",
    "أيوب مصري",
    "أسامة الوناس",
    "زيد كموش",
    "راشدي ياسين"
  ];
  saveGameData();
  res.json({ success: true, playerNames: gameState.playerNames });
});

// Upload Questions (From Excel or JSON)
app.post("/api/uploadQuestions", (req, res) => {
  gameState.questions = req.body.questions || [];
  gameState.status.totalQuestions = gameState.questions.length;
  if (gameState.status.currentQuestionId > (gameState.questions.length || 1)) {
    gameState.status.currentQuestionId = 1;
  }
  const q = gameState.questions[gameState.status.currentQuestionId - 1];
  gameState.status.correctAnswer = q ? q.correct : "";
  saveGameData();
  res.json({ success: true, count: gameState.questions.length });
});

// --- Start Server ---
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static("dist"));
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`🚀 MISHKAH Local Server running at http://localhost:${PORT}`);
    console.log(`📊 Local vMix Excel generated at: ${EXCEL_FILE}`);
  });
}

startServer();
